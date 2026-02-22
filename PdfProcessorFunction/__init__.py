import logging
import json
import os
import io
import fitz  # pymupdf
import azure.functions as func
from sharepoint_graph_utils import get_access_token, download_pdf_from_sharepoint, upload_pdf_to_sharepoint


def fill_pdf_fields(pdf_content, field_data):
    # PDF-Dokument aus Bytes öffnen
    doc = fitz.open(stream=pdf_content, filetype="pdf")
    
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
    logging.info('PDF processor function triggered')

    try:
        req_body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            "Invalid JSON in request body",
            status_code=400
        )

    source_folder = req_body.get('source_folder', '/Documents/Templates')
    source_filename = req_body.get('source_filename', 'template.pdf')
    dest_folder = req_body.get('dest_folder', '/Documents/Processed')
    dest_filename = req_body.get('dest_filename', 'filled_document.pdf')
    field_data = req_body.get('field_data', {})

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
        
        pdf_content = download_pdf_from_sharepoint(
            access_token, site_url, source_folder, source_filename
        )
        
        filled_pdf = fill_pdf_fields(pdf_content, field_data)
        
        upload_pdf_to_sharepoint(
            access_token, site_url, dest_folder, dest_filename, filled_pdf
        )
        
        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "message": f"PDF processed and saved to {dest_folder}/{dest_filename}"
            }),
            mimetype="application/json",
            status_code=200
        )
    
    except Exception as e:
        logging.error(f"Error processing PDF: {str(e)}")
        return func.HttpResponse(
            json.dumps({"status": "error", "message": str(e)}),
            mimetype="application/json",
            status_code=500
        )
