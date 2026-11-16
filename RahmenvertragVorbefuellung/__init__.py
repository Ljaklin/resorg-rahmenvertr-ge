import logging
import json
import os
import fitz  # pymupdf
import azure.functions as func
from sharepoint_graph_utils import (
    get_access_token, get_site_id, get_drive_by_name, get_item_web_url,
    copy_folder_recursive, sanitize_sharepoint_name, download_file_by_path, upload_file,
    get_list_items, resolve_sharepoint_url
)
from gaeb_utils import build_baustellen_bezeichnung, fill_x86_fields

# SharePoint list that maps a Bst.-Nr. to extra documents (e.g. the Vonovia LV).
# Columns (internal names): Title, Bst_x002e__x002d_Nr_x002e_ (Bst.-Nr.), Vorlage, Zielordner
TEMPLATE_CONFIG_LIST = os.getenv("template_config_list", "LV_Vorlagen_Rahmenverträge")
COL_TITLE = "Title"
COL_BST_NR = "Bst_x002e__x002d_Nr_x002e_"
COL_VORLAGE = "Vorlage"
COL_ZIELORDNER = "Zielordner"


def fill_document(filename, content, field_data):
    """Fill a template depending on its file type. Unknown types are copied unchanged."""
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".x86":
        bezeichnung = build_baustellen_bezeichnung(field_data)
        if not bezeichnung:
            raise ValueError("No address given (stadt/adresse or adresse_gesamt)")
        return fill_x86_fields(content, baustellenname=bezeichnung, lv_bezeichnung=bezeichnung)
    if ext == ".pdf":
        return fill_pdf_fields(field_data, content)
    return content


def create_configured_documents(access_token, site_id, field_data):
    """Create one document per row in the template config list whose Bst.-Nr. matches.
    Each file is saved as '{ang_nr}_{Titel}{ext}' in the row's Zielordner.
    Returns (created paths, error messages)."""
    bst_nr = str(field_data.get('bst_nr') or '').strip()
    if not bst_nr:
        return [], []

    rows = get_list_items(
        access_token, site_id, TEMPLATE_CONFIG_LIST,
        [COL_TITLE, COL_BST_NR, COL_VORLAGE, COL_ZIELORDNER]
    )
    rows = [r for r in rows if str(r.get(COL_BST_NR) or '').strip() == bst_nr]
    logging.info(f"{len(rows)} template config row(s) for Bst.-Nr. {bst_nr}")

    created, errors = [], []
    for row in rows:
        title = (row.get(COL_TITLE) or '').strip()
        try:
            ang_nr = str(field_data.get('ang_nr') or '').strip()
            if not ang_nr:
                raise ValueError("ang_nr is missing, cannot name the file")

            template_drive, template_path = resolve_sharepoint_url(access_token, site_id, row.get(COL_VORLAGE))
            target_drive, target_folder = resolve_sharepoint_url(access_token, site_id, row.get(COL_ZIELORDNER))

            template_name = template_path.rsplit("/", 1)[-1]
            ext = os.path.splitext(template_name)[1]
            filename = sanitize_sharepoint_name(f"{ang_nr}_{title}{ext}" if title else f"{ang_nr}{ext}")

            content = download_file_by_path(access_token, template_drive, template_path)
            filled = fill_document(template_name, content, field_data)
            upload_file(access_token, target_drive, target_folder, filename, filled)

            path = f"{target_folder}/{filename}" if target_folder else filename
            logging.info(f"Created '{path}' from template '{template_path}'")
            created.append(path)
        except Exception as e:
            error_msg = f"Error creating document '{title}': {str(e)}"
            logging.error(error_msg)
            errors.append(error_msg)

    return created, errors


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
    # Format: "{auftraggeber} - {bst_nr}/{stadt}, {adresse}, {adresszusatz}, {ersteller}, {ang_nr}"
    # (empty parts are skipped)
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

        # --- Extra documents per Bst.-Nr. from the template config list (e.g. Vonovia LV) ---
        try:
            documents, doc_errors = create_configured_documents(access_token, site_id, field_data)
        except Exception as e:
            documents, doc_errors = [], [f"Error reading template config list: {str(e)}"]
            logging.error(doc_errors[0])
        result["processed"].extend(documents)
        result["errors"].extend(doc_errors)

        # --- Folder URL for Power Automate, as reported by SharePoint ---
        folder_url = None
        try:
            folder_url = get_item_web_url(access_token, target_drive_id, dest_path)
            logging.info(f"Created folder URL: {folder_url}")
        except Exception as e:
            logging.warning(f"Could not resolve folder URL for {dest_path}: {str(e)}")

        response_data = {
            "status": "success" if result["processed"] else "error",
            "processed_files": len(result["processed"]),
            "files": result["processed"],
            "destination": dest_path,
            "folder_url": folder_url,
            "ang_nr": field_data.get('ang_nr'),
            "documents": documents
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
