import html
import logging
import requests
import re
from urllib.parse import urlparse, parse_qs, unquote

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def sanitize_sharepoint_name(name):
    """Sanitize a file or folder name for SharePoint compatibility.
    
    SharePoint does not allow these characters: " * : < > ? / \ |
    Also replaces leading/trailing spaces and periods.
    """
    if not name:
        return name
    
    # Replace invalid characters with safe alternatives
    replacements = {
        '/': '-',
        '\\': '-',
        ':': '-',
        '*': '-',
        '?': '',
        '"': "'",
        '<': '(',
        '>': ')',
        '|': '-'
    }
    
    for old_char, new_char in replacements.items():
        name = name.replace(old_char, new_char)
    
    # Remove leading/trailing spaces and periods
    name = name.strip('. ')
    
    # Replace multiple spaces with single space
    name = re.sub(r'\s+', ' ', name)
    
    return name


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


def list_drives(access_token, site_id):
    """List the document libraries (drives) of a site, including their webUrl."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    return response.json().get("value", [])


def get_drive_by_name(access_token, site_id, library_name):
    """Get a drive ID by document library display name (e.g. 'Templates', 'Documents')."""
    drives = list_drives(access_token, site_id)
    drive_names = [d["name"] for d in drives]
    logging.info(f"Available libraries on site: {drive_names}")
    # Case-insensitive match
    for drive in drives:
        if drive["name"].lower() == library_name.lower():
            return drive["id"]
    raise Exception(f"Document library '{library_name}' not found on site. Available: {drive_names}")


def sharepoint_url_to_server_path(url):
    """Extract the server-relative path from a SharePoint link as copied from the browser.

    Handles library view links ('.../Forms/AllItems.aspx?id=%2Fsites%2F...%2Ffile.x86'),
    direct file/folder links ('.../sites/X/Library/folder/file.x86') and links wrapped in
    rich-text HTML. Returns e.g. '/sites/Projektabwicklung/Templates/LV_Vonovia_Rahmenvertrag.x86'.
    """
    match = re.search(r'https?://[^\s"<>]+', html.unescape(url or ""))
    if not match:
        raise ValueError(f"No SharePoint URL found in '{url}'")
    parsed = urlparse(match.group(0))

    query = parse_qs(parsed.query)
    if "id" in query:
        server_path = query["id"][0]  # parse_qs already URL-decodes
    else:
        server_path = unquote(parsed.path)
        if "/Forms/" in server_path:  # library view without id = library root
            server_path = server_path.split("/Forms/")[0]
    return server_path.rstrip("/")


def resolve_sharepoint_url(access_token, site_id, url):
    """Resolve a SharePoint browser link to (drive_id, path inside the drive).
    The path is '' for the library root."""
    server_path = sharepoint_url_to_server_path(url)

    best = None
    for drive in list_drives(access_token, site_id):
        drive_path = unquote(urlparse(drive.get("webUrl", "")).path).rstrip("/")
        if not drive_path:
            continue
        if server_path.lower() == drive_path.lower() or server_path.lower().startswith(drive_path.lower() + "/"):
            if best is None or len(drive_path) > len(best[1]):
                best = (drive["id"], drive_path)

    if best is None:
        raise ValueError(f"No document library on this site contains '{server_path}'")
    drive_id, drive_path = best
    return drive_id, server_path[len(drive_path):].strip("/")


def get_list_items(access_token, site_id, list_name, field_names):
    """Return the field dicts of all items of a SharePoint list.
    list_name may be the display name or the URL name of the list (case-insensitive).
    field_names are internal column names, e.g. 'Bst_x002e__x002d_Nr_x002e_'."""
    url = f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,name,displayName"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    lists = response.json().get("value", [])
    list_id = next(
        (l["id"] for l in lists
         if list_name.lower() in (l.get("displayName", "").lower(), l.get("name", "").lower())),
        None
    )
    if list_id is None:
        raise ValueError(f"SharePoint list '{list_name}' not found on site")

    select = ",".join(field_names)
    url = f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?$expand=fields($select={select})&$top=200"
    items = []
    while url:
        response = requests.get(url, headers=_headers(access_token))
        response.raise_for_status()
        data = response.json()
        items.extend(item.get("fields", {}) for item in data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def get_item_web_url(access_token, drive_id, item_path):
    """Return the browser URL SharePoint reports for a file or folder path in a drive.
    Uses the library's real URL segment, which can differ from its display name."""
    encoded_path = requests.utils.quote(item_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}?$select=webUrl"
    response = requests.get(url, headers=_headers(access_token))
    response.raise_for_status()
    return response.json()["webUrl"]


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


def download_file_by_path(access_token, drive_id, file_path):
    """Download a file's content by its path relative to the drive root. Returns bytes."""
    encoded_path = requests.utils.quote(file_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/content"
    response = requests.get(url, headers=_headers(access_token), allow_redirects=True)
    response.raise_for_status()
    return response.content


def upload_file(access_token, drive_id, dest_folder_path, filename, content):
    """Upload a file (up to 4 MB) to a specific folder path in the drive."""
    # Sanitize the filename
    filename = sanitize_sharepoint_name(filename)
    file_path = f"{dest_folder_path.strip('/')}/{filename}" if dest_folder_path.strip('/') else filename
    encoded_path = requests.utils.quote(file_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/content"
    headers = _headers(access_token)
    headers["Content-Type"] = "application/octet-stream"
    response = requests.put(url, headers=headers, data=content)
    response.raise_for_status()
    return response.json()


def create_folder(access_token, drive_id, parent_path, folder_name):
    """Create a folder in the specified parent path. Creates the folder even if empty.
    
    parent_path: Path to the parent folder (e.g. 'Documents/Project')
    folder_name: Name of the new folder to create
    
    Returns the created folder item dict.
    """
    # Sanitize the folder name
    folder_name = sanitize_sharepoint_name(folder_name)
    encoded_path = requests.utils.quote(parent_path)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded_path}:/children"
    headers = _headers(access_token)
    headers["Content-Type"] = "application/json"
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace"
    }
    response = requests.post(url, headers=headers, json=body)
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
        sanitized_name = sanitize_sharepoint_name(item_name)
        
        if "folder" in item:
            # Explicitly create the folder in destination (handles empty folders too)
            try:
                logging.info(f"Creating folder: {dest_path}/{sanitized_name}")
                create_folder(access_token, dest_drive_id, dest_path, sanitized_name)
            except Exception as e:
                logging.warning(f"Folder creation note for {sanitized_name}: {str(e)}")
            
            # Recurse into subfolder
            logging.info(f"Entering subfolder: {sanitized_name}")
            sub_result = copy_folder_recursive(
                access_token, source_drive_id,
                f"{source_path}/{item_name}",
                dest_drive_id,
                f"{dest_path}/{sanitized_name}",
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
                
                logging.info(f"Uploading: {sanitized_name} -> {dest_path}/{sanitized_name}")
                upload_file(access_token, dest_drive_id, dest_path, sanitized_name, file_content)
                result["processed"].append(f"{dest_path}/{sanitized_name}")
                
            except Exception as e:
                error_msg = f"Error copying {item_name}: {str(e)}"
                logging.error(error_msg)
                result["errors"].append(error_msg)
    
    return result
