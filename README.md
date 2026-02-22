# resorg-rahmenvertr-ge

Azure Function for processing PDFs from SharePoint.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure environment variables in `local.settings.json`:
   - `tenant_id`: Azure AD tenant ID
   - `client_id`: Application (client) ID
   - `client_secret`: Client secret
   - `site_url`: SharePoint site URL
   - `resource`: SharePoint resource URL

## Usage

Send POST request to the function with JSON body:

```json
{
  "source_folder": "/Documents/Templates",
  "source_filename": "template.pdf",
  "dest_folder": "/Documents/Processed",
  "dest_filename": "filled_document.pdf",
  "field_data": {
    "name": "John Doe",
    "date": "2024-01-01",
    "company": "Acme Corp"
  }
}
```

The function will:
1. Download PDF from source SharePoint folder
2. Fill in the PDF form fields
3. Upload filled PDF to destination SharePoint folder