import logging
import json
import os
import fitz  # pymupdf
import azure.functions as func
from sharepoint_graph_utils import (
    get_access_token, get_site_id, get_drive_by_name, copy_folder_recursive
)


def fill_pdf_fields(field_data, pdf_content):
    """Fill PDF form fields from bytes. Returns processed PDF bytes."""
    doc = fitz.open(stream=pdf_content, filetype="pdf")
    
    for page in doc:
        for field in page.widgets():
            if field.field_name in field_data:
                logging.info(f"Feld gefunden: {field.field_name}, alter Wert: {field.field_value}")
                field.field_value = field_data[field.field_name]
                field.update()
                logging.info(f"Neuer Wert: {field_data[field.field_name]}")
    
    output = doc.tobytes()
    doc.close()
    return output


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Rahmenvertrag Vorbefuellung function triggered')

    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON in request body", status_code=400)

    # --- Extract field data from request ---
    field_data = {
        'strasse_hausnummer': req_body.get('strasse_hausnummer', ''),
        'plz_stadt': req_body.get('plz_stadt', ''),
        'adresse_gesamt': req_body.get('adresse_gesamt', ''),
        'bst_nr': req_body.get('bst_nr', ''),
        'ang_nr': req_body.get('ang_nr', ''),
        'Kombinationsfeld_Umkreis': req_body.get('Kombinationsfeld_Umkreis', ''),
        'gewerk': req_body.get('gewerk', ''),
        'auftraggeber': req_body.get('auftraggeber', ''),
        'datum': req_body.get('datum', ''),
        'ersteller': req_body.get('ersteller', ''),
    }

    # Optional fields from Personalplanung
    for optional_field in ('datum_unterweisung', 'ausführungszeit', 'mitarbeiter'):
        if req_body.get(optional_field):
            field_data[optional_field] = req_body[optional_field]

    # --- Environment variables ---
    tenant_id = os.getenv("tenant_id")
    client_id = os.getenv("client_id")
    client_secret = os.getenv("client_secret")
    site_url = os.getenv("site_url")
    resource = os.getenv("resource")
    source_library = os.getenv("source_library")  # SharePoint library name, e.g. "Templates"
    source_folder = os.getenv("source_folder")    # Path within library, e.g. "00.Musterordner Rahmenverträge"
    target_library = os.getenv("target_library")  # SharePoint library name, e.g. "Rahmenverträge"

    required_vars = {
        "tenant_id": tenant_id, "client_id": client_id, "client_secret": client_secret,
        "site_url": site_url, "resource": resource,
        "source_library": source_library, "source_folder": source_folder,
        "target_library": target_library
    }
    missing = [k for k, v in required_vars.items() if not v]
    if missing:
        return func.HttpResponse(
            json.dumps({"status": "error", "message": f"Missing env vars: {', '.join(missing)}"}),
            mimetype="application/json", status_code=500
        )

    # --- Build destination folder name ---
    # Format: "adresse_gesamt, bst_nr., ang_nr, ersteller, datum"
    auftraggeber = field_data.get('auftraggeber', 'Unbekannt')
    folder_name = ", ".join(filter(None, [
        field_data.get('adresse_gesamt'),
        field_data.get('bst_nr'),
        field_data.get('ang_nr'),
        field_data.get('ersteller'),
        field_data.get('datum'),
    ]))
    if not folder_name:
        folder_name = "Unbenannt"

    # Output goes into: {auftraggeber}/{folder_name} at the root of the target library
    dest_path = f"{auftraggeber}/{folder_name}"
    logging.info(f"Source: {source_library}/{source_folder}")
    logging.info(f"Destination: {target_library}/{dest_path}")

    try:
        # --- Authenticate and resolve IDs ---
        access_token = get_access_token(tenant_id, client_id, client_secret, resource)
        site_id = get_site_id(access_token, site_url)
        source_drive_id = get_drive_by_name(access_token, site_id, source_library)
        target_drive_id = get_drive_by_name(access_token, site_id, target_library)

        # --- Create a PDF processor closure with the field data ---
        def pdf_processor(pdf_bytes):
            return fill_pdf_fields(field_data, pdf_bytes)

        # --- Copy entire folder, filling PDFs along the way ---
        result = copy_folder_recursive(
            access_token, source_drive_id,
            source_folder, target_drive_id, dest_path,
            pdf_processor=pdf_processor
        )

        response_data = {
            "status": "success" if result["processed"] else "error",
            "processed_files": len(result["processed"]),
            "files": result["processed"],
            "destination": dest_path,
        }
        if result["errors"]:
            response_data["errors"] = result["errors"]
            response_data["status"] = "partial" if result["processed"] else "error"

        return func.HttpResponse(
            json.dumps(response_data),
            mimetype="application/json",
            status_code=200 if result["processed"] else 500
        )

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return func.HttpResponse(
            json.dumps({"status": "error", "message": str(e)}),
            mimetype="application/json", status_code=500
        )
