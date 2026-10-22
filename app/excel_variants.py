"""Select sensor-count templates from their saved workbook input."""

from decimal import Decimal, InvalidOperation
import re

from app.pdf import CertificateError


SENSOR_TEMPLATE = re.compile(r"(?:[0-9]+[- _])?certificado[_ -]([0-9]{1,3})s", re.I)


def sensor_template(parts, workbook, targets, strings):
    from app.excel import NS, R, xml, cell_value

    variants = [(index, entry, SENSOR_TEMPLATE.fullmatch(entry.get("name", "").strip()))
                for index, entry in enumerate(workbook.findall("s:sheets/s:sheet", NS))]
    variants = [(index, entry, int(match[1])) for index, entry, match in variants
                if match and entry.get("state", "visible") == "visible"]
    if not variants:
        return None, ""
    def fallback():
        view = workbook.find('s:bookViews/s:workbookView',NS)
        active = int(view.get('activeTab','0')) if view is not None else 0
        index,entry,_ = next((item for item in variants if item[0]==active),variants[0])
        return index,f'El registro no indica una cantidad de sensores; se eligió un solo formato: {entry.get("name")}.'
    names = [name for name in workbook.findall("s:definedNames/s:definedName", NS)
             if name.get("name", "").casefold() == "cantsensores" and name.get("localSheetId") is None]
    reference = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^'!\[\]]+))!\$?([A-Z]{1,3})\$?([1-9][0-9]*)",
                             (names[0].text or "").strip(), re.I) if len(names) == 1 else None
    if reference is None:
        if not names:
            return fallback()
        raise CertificateError("La referencia CantSensores del registro no es válida. Revisa ese dato para elegir el certificado correcto.")
    sheet_name = reference[1].replace("''", "'") if reference[1] else reference[2]
    source = next((entry for entry in workbook.findall("s:sheets/s:sheet", NS)
                   if entry.get("name", "").casefold() == sheet_name.casefold()), None)
    path = targets.get(source.get(f"{{{R}}}id")) if source is not None else None
    cell_ref = reference[3].upper() + reference[4]
    cell = xml(parts[path]).find(f"s:sheetData/s:row/s:c[@r='{cell_ref}']", NS) if path in parts else None
    if cell is None or not cell_value(cell,strings).strip():
        return fallback()
    try:
        value = Decimal(cell_value(cell, strings)) if cell is not None and cell.get("t") != "e" else Decimal("NaN")
        chosen = [(index, entry) for index, entry, count in variants if value == count]
    except (InvalidOperation, ValueError):
        chosen = []
    if len(chosen) != 1:
        raise CertificateError("La cantidad de sensores del registro no corresponde a un único formato de certificado. Revisa CantSensores y guarda el Excel.")
    index, entry = chosen[0]
    return index, f"Se eligió {entry.get('name')} porque el registro {sheet_name}!{cell_ref} indica {int(value)} sensores."
