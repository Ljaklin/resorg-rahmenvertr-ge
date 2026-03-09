import logging
import json
import os
import io
import fitz  # pymupdf
import azure.functions as func
from sharepoint_graph_utils import get_access_token, upload_pdf_to_sharepoint


# Template files to process
TEMPLATE_FILES = [
    "Betriebsanweisung, Arbeitsplan AZ.PDF",
    "Betriebsanweisung, Arbeitsplan Flexplatten.PDF",
    "Deckblatt, Gefährdungsbeurteilung und Aufmaß.pdf",
    "Meldung AZ.pdf",
    "Meldung Flexplatten.pdf"
]


def fill_pdf_fields(pdf_path, field_data):
    """Fill PDF form fields from a local template file."""
    doc = fitz.open(pdf_path)
    
    # Formularfelder bearbeiten
    for page in doc:
        for field in page.widgets():  # Alle Formularfelder auf der Seite durchlaufen
            if field.field_name in field_data:  # Prüfen, ob das Feld aktualisiert werden soll
                logging.info(f"Feld gefunden: {field.field_name}, alter Wert: {field.field_value}")
                field.field_value = field_data[field.field_name]  # Neuen Wert setzen
                field.update()  # Feld aktualisieren
                logging.info(f"Neuer Wert: {field_data[field.field_name]}")
    
    # Geändertes PDF als Bytes zurückgeben
    output = doc.tobytes()
    doc.close()
    return output


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Rahmenvertrag Vorbefüllung function triggered')

    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            "Invalid JSON in request body",
            status_code=400
        )

    # Extract field data from request
    # Canvas sends: strasse_hausnummer, plz_stadt, adresse_gesamt, bst_nr, ang_nr,
    # Kombinationsfeld_Umkreis, gewerk, auftraggeber, datum, ersteller
    # Optional: datum_unterweisung, ausführungszeit, mitarbeiter
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
    if req_body.get('datum_unterweisung'):
        field_data['datum_unterweisung'] = req_body.get('datum_unterweisung')
    if req_body.get('ausführungszeit'):
        field_data['ausführungszeit'] = req_body.get('ausführungszeit')
    if req_body.get('mitarbeiter'):
        field_data['mitarbeiter'] = req_body.get('mitarbeiter')

    # SharePoint destination folder
    dest_folder = req_body.get('dest_folder', '/Documents/Rahmenverträge')
    
    # Get SharePoint credentials
    tenant_id = os.getenv("tenant_id")
    client_id = os.getenv("client_id")
    client_secret = os.getenv("client_secret")
    site_url = os.getenv("site_url")
    resource = os.getenv("resource")

    if not all([tenant_id, client_id, client_secret, site_url, resource]):
        return func.HttpResponse(
            "Missing required environment variables",
            status_code=500
        )

    try:
        access_token = get_access_token(tenant_id, client_id, client_secret, resource)
        
        # Get the base directory of the function
        function_dir = os.path.dirname(os.path.abspath(__file__))
        template_dir = os.path.join(os.path.dirname(function_dir), "data", "00.Musterordner Rahmenverträge")
        
        if not os.path.exists(template_dir):
            raise Exception(f"Template directory not found: {template_dir}")
        
        processed_files = []
        errors = []
        
        # Process each template file
        for template_file in TEMPLATE_FILES:
            try:
                template_path = os.path.join(template_dir, template_file)
                
                if not os.path.exists(template_path):
                    logging.warning(f"Template file not found: {template_path}")
                    errors.append(f"Template not found: {template_file}")
                    continue
                
                logging.info(f"Processing template: {template_file}")
                
                # Fill the PDF with field data
                filled_pdf = fill_pdf_fields(template_path, field_data)
                
                # Generate destination filename using bst_nr or ang_nr as prefix
                prefix = field_data.get('bst_nr', field_data.get('ang_nr', 'document'))
                dest_filename = f"{prefix}_{template_file}"
                
                # Upload to SharePoint
                upload_pdf_to_sharepoint(
                    access_token, site_url, dest_folder, dest_filename, filled_pdf
                )
                
                processed_files.append(dest_filename)
                logging.info(f"Successfully processed: {template_file}")
                
            except Exception as e:
                error_msg = f"Error processing {template_file}: {str(e)}"
                logging.error(error_msg)
                errors.append(error_msg)
        
        # Prepare response
        response_data = {
            "status": "success" if processed_files else "error",
            "processed_files": len(processed_files),
            "files": processed_files,
            "destination": dest_folder
        }
        
        if errors:
            response_data["errors"] = errors
            response_data["status"] = "partial" if processed_files else "error"
        
        status_code = 200 if processed_files else 500
        
        return func.HttpResponse(
            json.dumps(response_data),
            mimetype="application/json",
            status_code=status_code
        )
    
    except Exception as e:
        logging.error(f"Error processing PDFs: {str(e)}")
        return func.HttpResponse(
            json.dumps({"status": "error", "message": str(e)}),
            mimetype="application/json",
            status_code=500
        )
