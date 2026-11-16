# resorg-rahmenvertr-ge

Azure Function (`RahmenvertragVorbefuellung`) that provisions a new project folder for a Rahmenvertrag in SharePoint.
It copies a template folder from one document library into another and pre-fills the form fields of every PDF along the way.

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure environment variables (`local.settings.json` locally, App Settings in Azure):

   | Variable | Description |
   |---|---|
   | `tenant_id` | Entra ID tenant ID |
   | `client_id` | App registration (client) ID |
   | `client_secret` | App registration client secret |
   | `site_url` | SharePoint site in Graph form, e.g. `contoso.sharepoint.com:/sites/Projektabwicklung` |
   | `resource` | Token scope prefix, e.g. `https://graph.microsoft.com/` |
   | `source_library` | Display name of the library holding the template folder, e.g. `Templates` |
   | `source_folder` | Path of the template folder inside `source_library`, e.g. `00.Musterordner Rahmenverträge` |
   | `target_library` | Display name of the library the new project folder is created in, e.g. `Rahmenverträge` |

   Optional:

   | Variable | Default | Description |
   |---|---|---|
   | `template_config_list` | `LV_Vorlagen_Rahmenverträge` | SharePoint list (display or URL name) on the same site that maps a Bst.-Nr. to extra documents, see below |

   Library names are matched case-insensitively. The app registration needs Graph application permissions to read the source library and write to the target library (e.g. `Sites.ReadWrite.All` or `Sites.Selected`).

## Usage

Send a `POST` request (function key required) with the project data as JSON:

```json
{
  "auftraggeber": "Musterbau GmbH",
  "bst_nr": "4711",
  "ang_nr": "A-2024-001",
  "stadt": "Köln",
  "plz": "50667",
  "adresse": "Domstraße 1",
  "adresszusatz": "Hinterhaus",
  "strasse_hausnummer": "Domstraße 1",
  "plz_stadt": "50667 Köln",
  "adresse_gesamt": "Domstraße 1, 50667 Köln",
  "gewerk": "Abbruch",
  "ersteller": "MM",
  "datum": "01.01.2024",
  "Kombinationsfeld_Umkreis": "",
  "datum_unterweisung": "02.01.2024",
  "ausführungszeit": "KW 2-4",
  "mitarbeiter": "Max Mustermann"
}
```

All fields are optional; missing ones are filled with an empty string. `datum_unterweisung`, `ausführungszeit` and `mitarbeiter` (from Personalplanung) are only written to the PDFs when present.

The function will:
1. Build the destination path `{auftraggeber} - {bst_nr}/{stadt}, {adresse}, {adresszusatz}, {ersteller}, {ang_nr}` (empty parts are skipped, characters SharePoint does not allow are replaced).
2. Copy `source_folder` recursively into that path in `target_library`.
3. Fill every PDF form field whose name matches a key of the request body. Non-PDF files are copied unchanged.
4. Create the extra documents configured for the request's `bst_nr` in the template config list (see below).
5. Return a JSON summary:

```json
{
  "status": "success",
  "processed_files": 5,
  "files": ["Musterbau GmbH - 4711/Köln, Domstraße 1, Hinterhaus, MM, A-2024-001/Meldung AZ.pdf"],
  "destination": "Musterbau GmbH - 4711/Köln, Domstraße 1, Hinterhaus, MM, A-2024-001",
  "folder_url": "https://contoso.sharepoint.com/sites/Projektabwicklung/...",
  "ang_nr": "A-2024-001",
  "documents": ["Vonovia LVs zum Anlegen/A-2024-001_Vonovia.x86"]
}
```

`documents` lists the extra documents created from the template config list (empty if none are configured for the Bst.-Nr.; failures are listed in `errors`).

## Template config list (extra documents per Bst.-Nr.)

The list `LV_Vorlagen_Rahmenverträge` on the Projektabwicklung site controls which extra documents are created, without a code change or deployment:

| Column | Internal name | Content |
|---|---|---|
| Titel | `Title` | Short name, used in the file name, e.g. `Vonovia` |
| Bst.-Nr. | `Bst_x002e__x002d_Nr_x002e_` | Bst.-Nr. this row applies to, e.g. `730080` |
| Vorlage | `Vorlage` | Link to the template file, copied from the browser |
| Zielordner | `Zielordner` | Link to the target folder, copied from the browser |

For every row whose Bst.-Nr. matches the request, the function downloads the template, fills it and saves it as
`{ang_nr}_{Titel}{extension}` (e.g. `26-1865_Vonovia.x86`) in the target folder. A file with the same name is overwritten.
Several rows with the same Bst.-Nr. create several documents.

How a template is filled depends on its file extension:

| Extension | Filled with |
|---|---|
| `.x86` (GAEB DA XML 3.2) | *Baustellenname* (`PrjInfo/LblPrj`) and *LV-Bezeichnung* (`BoQ/BoQInfo/LblBoQ`) = `Stadt (PLZ), Straße, Adresszusatz`, built from `stadt`, `plz`, `adresse`, `adresszusatz`; if `stadt` or `adresse` is missing, `adresse_gesamt` is used as-is |
| `.pdf` | Form fields whose name matches a request key |
| anything else | Copied unchanged |

Template and target folder may be in any document library of the site. Both links are accepted as library view links (`.../Forms/AllItems.aspx?id=...`) or direct links. A new file type or new field mapping needs a code change in `fill_document` (`RahmenvertragVorbefuellung/__init__.py`).

`status` is `success`, `partial` (some files failed, listed in `errors`) or `error`. Existing folders with the same name are replaced.

## Tools

- `inspect_pdf_fields.py` prints the form fields (name, type, current value) of every PDF in `data/00.Musterordner Rahmenverträge`, to check which request keys the templates expect. Run it from the repo root.
