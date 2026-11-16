import xml.etree.ElementTree as ET

GAEB_NS = "http://www.gaeb.de/GAEB_DA_XML/DA86/3.2"


def build_baustellen_bezeichnung(field_data):
    """Build 'Stadt (PLZ), Straße, Adresszusatz' for GAEB Baustellenname / LV-Bezeichnung.

    Uses the single address parts when stadt and street are given. Otherwise falls back to
    'adresse_gesamt', which already arrives in that format. Returns '' if neither is available.
    """
    stadt = (field_data.get('stadt') or '').strip()
    plz = (field_data.get('plz') or '').strip()
    strasse = (field_data.get('adresse') or field_data.get('strasse_hausnummer') or '').strip()
    zusatz = (field_data.get('adresszusatz') or '').strip()

    if stadt and strasse:
        ort = f"{stadt} ({plz})" if plz else stadt
        return ", ".join(filter(None, [ort, strasse, zusatz]))

    return (field_data.get('adresse_gesamt') or '').strip()


def fill_x86_fields(x86_content, baustellenname, lv_bezeichnung):
    """Set Baustellenname (PrjInfo/LblPrj) and LV-Bezeichnung (BoQ/BoQInfo/LblBoQ)
    in a GAEB DA XML 3.2 file (DA86). Returns the new file content as UTF-8 bytes."""
    ns = {"g": GAEB_NS}
    ET.register_namespace("", GAEB_NS)
    root = ET.fromstring(x86_content)

    targets = {
        "g:PrjInfo/g:LblPrj": baustellenname,
        "g:Award/g:BoQ/g:BoQInfo/g:LblBoQ": lv_bezeichnung,
    }
    for path, value in targets.items():
        element = root.find(path, ns)
        if element is None:
            raise ValueError(f"GAEB element '{path}' not found in template")
        element.text = value

    return ET.tostring(root, encoding="UTF-8", xml_declaration=True) + b"\n"
