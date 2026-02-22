import os
import requests


def get_access_token(tenant_id, client_id, client_secret, resource):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "resource": resource
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]


def download_pdf_from_sharepoint(access_token, site_url, folder_path, filename):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose"
    }
    file_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{folder_path}/{filename}')/$value"
    response = requests.get(file_url, headers=headers)
    response.raise_for_status()
    return response.content


def upload_pdf_to_sharepoint(access_token, site_url, folder_path, filename, pdf_content):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/pdf"
    }
    upload_url = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{folder_path}')/Files/add(url='{filename}',overwrite=true)"
    response = requests.post(upload_url, headers=headers, data=pdf_content)
    response.raise_for_status()
    return response.json()
