import logging
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token(tenant_id, client_id, client_secret, resource):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": f"{resource}.default"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    token = response.json().get("access_token")
    if not token:
        raise Exception("Failed to retrieve access token from response")
    return token


def _headers(access_token):
    return {"Authorization": f"Bearer {access_token}"}


def get_site_id(access_token, site_url):
    """Resolve a SharePoint site URL (e.g. 'contoso.sharepoint.com:/sites/MySite') to a site ID."""
    url = f"{GRAPH_BASE}/sites/{site_url}"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    return response.json()["id"]


def get_drive_id(access_token, site_id):
    """Get the default document library drive ID for a site."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    return response.json()["id"]


def get_drive_by_name(access_token, site_id, library_name):
    """Get a drive ID by document library display name (e.g. 'Templates', 'Documents')."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    drives = response.json().get("value", [])
    drive_names = [d["name"] for d in drives]
    logging.info(f"Available libraries on site: {drive_names}")
    # Case-insensitive match
    for drive in drives:
        if drive["name"].lower() == library_name.lower():
            return drive["id"]
    raise Exception(f"Document library '{library_name}' not found on site. Available: {drive_names}")


def list_folder_children(access_token, drive_id, folder_path):
    """List all children (files and folders) in a folder path.
    folder_path should be relative to the drive root, e.g. 'Musterordner/Vorlagen'.
    Returns a list of item dicts with 'name', 'id', 'file'/'folder' keys."""
    encoded_path = requests.utils.quote(folder_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/children"
    items = []
    while url:
        response = requests.get(url, headers=_headers(access_token))
        response.raise_for_status()
        data = response.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def download_file(access_token, drive_id, item_id):
    """Download a file's content by its item ID. Returns bytes."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    response = requests.get(url, headers=_headers(access_token), allow_redirects=True)
    response.raise_for_status()
    return response.content


def upload_file(access_token, drive_id, dest_folder_path, filename, content):
    """Upload a file (up to 4 MB) to a specific folder path in the drive."""
    encoded_path = requests.utils.quote(f"{dest_folder_path}/{filename}")
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/content"
    headers = _headers(access_token)
    headers["Content-Type"] = "application/octet-stream"
    response = requests.put(url, headers=headers, data=content)
    response.raise_for_status()
    return response.json()


def copy_folder_recursive(access_token, source_drive_id, source_path, dest_drive_id, dest_path, pdf_processor=None):
    """Recursively copy a folder from source to destination (can be different drives/libraries).
    
    If pdf_processor is provided, it will be called for each PDF file:
        pdf_processor(pdf_bytes) -> processed_pdf_bytes
    Non-PDF files are copied as-is.
    
    Returns a dict with 'processed' and 'errors' lists.
    """
    result = {"processed": [], "errors": []}
    
    children = list_folder_children(access_token, source_drive_id, source_path)
    
    for item in children:
        item_name = item["name"]
        
        if "folder" in item:
            # Recurse into subfolder
            logging.info(f"Entering subfolder: {item_name}")
            sub_result = copy_folder_recursive(
                access_token, source_drive_id,
                f"{source_path}/{item_name}",
                dest_drive_id,
                f"{dest_path}/{item_name}",
                pdf_processor=pdf_processor
            )
            result["processed"].extend(sub_result["processed"])
            result["errors"].extend(sub_result["errors"])
        
        elif "file" in item:
            try:
                logging.info(f"Downloading: {item_name}")
                file_content = download_file(access_token, source_drive_id, item["id"])
                
                # Process PDFs if a processor is provided
                if pdf_processor and item_name.lower().endswith(".pdf"):
                    logging.info(f"Processing PDF: {item_name}")
                    file_content = pdf_processor(file_content)
                
                logging.info(f"Uploading: {item_name} -> {dest_path}/{item_name}")
                upload_file(access_token, dest_drive_id, dest_path, item_name, file_content)
                result["processed"].append(f"{dest_path}/{item_name}")
                
            except Exception as e:
                error_msg = f"Error copying {item_name}: {str(e)}"
                logging.error(error_msg)
                result["errors"].append(error_msg)
    
    return result
