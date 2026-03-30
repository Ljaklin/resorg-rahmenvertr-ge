import logging
import json
import os
import fitz  # pymupdf
import azure.functions as func
from sharepoint_graph_utils import (
    get_access_token, get_site_id, get_drive_by_name, copy_folder_recursive, sanitize_sharepoint_name
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
        'plz': req_body.get('plz', ''),
        'stadt': req_body.get('stadt', ''),
        'adresse': req_body.get('adresse', ''),
        'adresszusatz': req_body.get('adresszusatz', ''),
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
    # Format: "plz, stadt, adresse, adresszusatz, bst.-nr., auftraggeber, ersteller, datum, ang.-nr."
    auftraggeber = field_data.get('auftraggeber', 'Unbekannt')
    bst_nr = field_data.get('bst_nr', '')
    
    # Client folder includes Bst.-Nr.: "Client - Bst.-Nr."
    if bst_nr:
        client_folder = f"{auftraggeber} - {bst_nr}"
    else:
        client_folder = auftraggeber
    
    # Sanitize client folder name
    client_folder = sanitize_sharepoint_name(client_folder)
    
    folder_name = ", ".join(filter(None, [
        field_data.get('stadt'),
        field_data.get('adresse'),
        field_data.get('adresszusatz'),
        field_data.get('ersteller'),
        field_data.get('ang_nr'),
    ]))
    if not folder_name:
        folder_name = "Unbenannt"
    
    # Sanitize folder name
    folder_name = sanitize_sharepoint_name(folder_name)

    # Output goes into: {client_folder}/{folder_name} at the root of the target library
    dest_path = f"{client_folder}/{folder_name}"
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

        # --- Construct folder URL for Power Automate ---
        from urllib.parse import quote
        site_path = site_url.split(":/")[1]  # Extract "sites/Projektabwicklung"
        domain = site_url.split(":/")[0]  # Extract "118016aplus.sharepoint.com"
        
        # URL-encode the path, keeping forward slashes
        encoded_path = quote(dest_path, safe="/")
        folder_url = f"https://{domain}/{site_path}/{target_library}/{encoded_path}"
        logging.info(f"Created folder URL: {folder_url}")

        response_data = {
            "status": "success" if result["processed"] else "error",
            "processed_files": len(result["processed"]),
            "files": result["processed"],
            "destination": dest_path,
            "folder_url": folder_url,
            "ang_nr": field_data.get('ang_nr')
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
