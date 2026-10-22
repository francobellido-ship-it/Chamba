"""Export the saved values and print layout of a QCP XLSX certificate."""

from dataclasses import dataclass, replace
from io import BytesIO
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import subprocess
import tempfile
from time import monotonic
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image
import pymupdf

from app.pdf import CertificateError, MAX_PAGES
from app.backgrounds import Letterhead, apply_letterhead
from app.excel_headers import header_pictures, restore_header_pictures
from app.excel_package import print_package
from app.excel_ranges import printed_cell_checker, print_names
from app.excel_pages import fit_stationery_content
from app.excel_auxiliary import omit_auxiliary_errors
from app.excel_print_width import background_width_overflows
from app.excel_footers import FormatFooter, take_format_footers, footer_for_page, apply_format_footers, general_data_footer, entered_footer
from app.pdf_images import compact_backgrounds
from app.excel_rows import fit_table_headings, fit_table_fonts, wrap_table_lines, fit_overflow_labels
from app.excel_composition import compose_numbered_pages, drawing_row_ranges, validate_page_labels, validate_table_text, fit_manual_numbered_pages, validate_saved_labels

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"s": S, "r": R, "p": P, "d": D, "a": A}
ET.register_namespace("", S)
ET.register_namespace("r", R)


@dataclass
class ConvertedExcel:
    pdf: bytes
    embedded_signature: bool
    warnings: list[str]
    background_digests: tuple[bytes, ...] = ()
    letterhead_pages: tuple[int, ...] | None = None
    format_footers: tuple[FormatFooter | None, ...] = ()


@dataclass(frozen=True)
class ExcelLayout:
    sheet_name: str
    sheet_path: str
    sheet_index: int
    preserve_format: bool
    needs_background: bool = False
    page_numbers: tuple[int, ...] = ()
    selection_reason: str = ""


def _read_parts(data: bytes) -> dict[str, bytes]:
    try:
        source = zipfile.ZipFile(BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise CertificateError("El archivo debe ser un Excel .xlsx o .xlsm válido y sin contraseña.") from exc
    with source:
        entries = source.infolist()
        if len(entries) > 2000 or sum(e.file_size for e in entries) > 80 * 1024 * 1024:
            raise CertificateError("El contenido del Excel supera el tamaño permitido.")
        if len({e.filename for e in entries}) != len(entries):
            raise CertificateError("El Excel contiene partes duplicadas.")
        if any(e.filename.startswith("/") or ".." in PurePosixPath(e.filename).parts for e in entries):
            raise CertificateError("El Excel contiene rutas internas inválidas.")
        if any(e.flag_bits & 1 for e in entries):
            raise CertificateError("El Excel no debe tener contraseña.")
        parts = {entry.filename: source.read(entry) for entry in entries}
    if "xl/workbook.xml" not in parts:
        raise CertificateError("El archivo no contiene un libro Excel válido.")
    return parts


def _layout(parts: dict[str, bytes], sheet_index: int | None = None) -> ExcelLayout:
    workbook = xml(parts["xl/workbook.xml"])
    rels = xml(parts["xl/_rels/workbook.xml.rels"])
    targets = {rel.get("Id"): related("xl/workbook.xml", rel.get("Target", ""))
               for rel in rels if rel.get("TargetMode") != "External"}
    strings = (["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
               if "xl/sharedStrings.xml" in parts else [])
    reason = ""
    if sheet_index is None:
        from app.excel_variants import sensor_template
        sheet_index, reason = sensor_template(parts, workbook, targets, strings)
    candidates = []
    for index, entry in enumerate(workbook.findall("s:sheets/s:sheet", NS)):
        name = entry.get("name", "")
        if (sheet_index is not None and index != sheet_index) or (sheet_index is None
                and (entry.get('state','visible') != 'visible'
                     or not re.fullmatch(r"certificado(?:\s+qcp)?(?:\s*[0-9]+)?", name.strip(), re.I))):
            continue
        path = targets.get(entry.get(f"{{{R}}}id"))
        if not path or path not in parts:
            raise CertificateError("No se pudo localizar la hoja del certificado.")
        sheet = xml(parts[path])
        text = " ".join(cell_value(cell, strings) for cell in sheet.findall("s:sheetData/s:row/s:c", NS))
        marker = bool(re.search(r"escane\w*\s+(?:este\s+)?qr|c[oó]digo\s+qr", text, re.I))
        candidates.append((marker, index, name, path, sheet))
    if not candidates:
        raise CertificateError("El Excel debe tener una hoja llamada Certificado o CERTIFICADO QCP.")
    plain = [candidate for candidate in candidates if re.fullmatch(r'certificado(?:\s+qcp)?',candidate[2].strip(),re.I)]
    choices = plain or candidates
    marked = [candidate for candidate in choices if candidate[0]]
    qcp = [candidate for candidate in choices if re.fullmatch(r'certificado\s+qcp',candidate[2].strip(),re.I)]
    if len(plain)>1 and qcp:
        chosen=qcp[0]
        reason=f'Se eligió {chosen[2]} como único formato principal del libro; se conservarán sus páginas continuadas.'
    elif len(marked) == 1:
        chosen = marked[0]
    else:
        first = [candidate for candidate in choices if any(re.fullmatch(r'p[aá]gina\s*1\s*de\s*[0-9]+',cell_value(cell,strings).strip(),re.I)
                 for cell in candidate[4].findall('s:sheetData/s:row/s:c',NS))]
        view = workbook.find('s:bookViews/s:workbookView',NS)
        active = int(view.get('activeTab','0')) if view is not None else 0
        chosen = (qcp or (first if len(first)==1 else []) or [candidate for candidate in choices if candidate[1]==active] or choices)[0]
        if len(plain)>1:
            reason = f'Se eligió {chosen[2]} como único formato principal del libro; se conservarán sus páginas continuadas.'
    _, index, name, path, sheet = chosen
    picture = sheet.find("s:picture", NS)
    if picture is not None:
        rel_path = relationship_file(path)
        rels = xml(parts.get(rel_path, b"<Relationships/>"))
        target = next((rel.get("Target") for rel in rels if rel.get("Id") == picture.get(f"{{{R}}}id")), None)
        if not target or related(path, target) not in parts:
            raise CertificateError("Falta la imagen del fondo dentro del Excel. Guarda una copia con el fondo incorporado.")
    header = sheet.find("s:headerFooter", NS)
    existing = picture is not None or (header is not None and any((e.text or "").strip() for e in header))
    existing |= sheet.find("s:legacyDrawingHF", NS) is not None
    # Logos/letterheads drawn in worksheet cells must also remain intact.
    drawing = sheet.find("s:drawing", NS)
    if drawing is not None:
        rel_path = relationship_file(path)
        if rel_path in parts:
            drawing_rels = xml(parts[rel_path])
            target = next((rel.get("Target") for rel in drawing_rels
                           if rel.get("Id") == drawing.get(f"{{{R}}}id")), None)
            if target:
                root = xml(parts[related(path, target)])
                for anchor in root:
                    row = anchor.find("d:from/d:row", NS)
                    if anchor.find(".//a:blip", NS) is not None and row is not None and int(row.text) < 5:
                        existing = True
    return ExcelLayout(name, path, index, bool(existing), not bool(existing), selection_reason=reason)


def inspect_excel(data: bytes) -> ExcelLayout:
    try:
        from app.excel_sheets import certificate_layouts
        layouts = certificate_layouts(data)
        return replace(layouts[0], needs_background=any(layout.needs_background for layout in layouts))
    except CertificateError:
        raise
    except (KeyError, IndexError, ValueError, OSError, AttributeError, TypeError, zipfile.BadZipFile) as exc:
        raise CertificateError("No se pudo leer la estructura de este Excel.") from exc


def xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise CertificateError("El Excel contiene XML no permitido.")
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise CertificateError("El Excel está dañado.") from exc


def related(base: str, target: str) -> str:
    if ":" in target or target.startswith("//"):
        raise CertificateError("El Excel contiene una referencia externa no permitida.")
    parts = [] if target.startswith("/") else list(PurePosixPath(base).parent.parts)
    for part in target.split("/"):
        if part == "..":
            if not parts:
                raise CertificateError("Referencia inválida dentro del Excel.")
            parts.pop()
        elif part not in {"", "."}:
            parts.append(part)
    return "/".join(parts)


def relationship_file(part: str) -> str:
    path = PurePosixPath(part)
    return str(path.parent / "_rels" / (path.name + ".rels"))


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find("s:v", NS)
    if cell.get("t") == "s" and value is not None:
        return strings[int(value.text or "0")]
    if cell.get("t") == "inlineStr":
        inline = cell.find("s:is", NS)
        return "".join(inline.itertext()) if inline is not None else ""
    return value.text or "" if value is not None else ""


def _number_formats(data: bytes) -> bytes:
    """Match the reference's decimal comma and space between thousands."""
    styles = xml(data)
    formats = styles.find("s:numFmts", NS)
    if formats is None:
        formats = ET.Element(f"{{{S}}}numFmts")
        styles.insert(0, formats)
    # OOXML locale 040C gives decimal comma and a nonbreaking thousands space.
    # Keep dates and explicitly localized/currency formats as saved in the file.
    for fmt in formats:
        code = fmt.get("formatCode", "")
        plain = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", code)
        if re.search(r"[0#?]", plain) and not re.search(r"[ymdhs]", plain, re.I) and "[$" not in code:
            fmt.set("formatCode", "[$-040C]" + code)
    builtin = {1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
               9: "0%", 10: "0.00%", 11: "0.00E+00"}
    next_id = max([163] + [int(fmt.get("numFmtId")) for fmt in formats]) + 1
    mapped = {}
    for xf in styles.findall("s:cellXfs/s:xf", NS):
        old = int(xf.get("numFmtId", "0"))
        if old in builtin:
            if old not in mapped:
                mapped[old] = next_id
                ET.SubElement(formats, f"{{{S}}}numFmt", numFmtId=str(next_id), formatCode="[$-040C]" + builtin[old])
                next_id += 1
            xf.set("numFmtId", str(mapped[old]))
    formats.set("count", str(len(formats)))
    return ET.tostring(styles, encoding="utf-8", xml_declaration=True)


def _prepare_workbook(data: bytes, replace_signature: bool,
                      letterhead: Letterhead | None = None, headers: tuple | None = None,
                      layout: ExcelLayout | None = None) -> tuple[bytes, bytes | None, str, bool, list[str]]:
    parts = _read_parts(data)
    layout = layout or _layout(parts)
    if layout.preserve_format:
        letterhead = None
    # Export saved results offline. Data connections are not needed for printing.
    discarded = {name for name in parts if name.startswith(("xl/externalLinks/", "xl/queryTables/"))
                 or name == "xl/connections.xml"}
    for name in discarded:
        del parts[name]
    for name, content in list(parts.items()):
        if name.endswith(".rels"):
            root = xml(content)
            for rel in list(root):
                if rel.get("TargetMode") == "External" or rel.get("Type", "").rsplit("/", 1)[-1] in {"connections", "externalLink", "queryTable"}:
                    root.remove(rel)
            parts[name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    content_types = xml(parts["[Content_Types].xml"])
    for entry in list(content_types):
        if entry.get("PartName", "").lstrip("/") in discarded:
            content_types.remove(entry)
    parts["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    workbook = xml(parts["xl/workbook.xml"])
    parts["xl/styles.xml"] = _number_formats(parts["xl/styles.xml"])
    external = workbook.find("s:externalReferences", NS)
    if external is not None:
        workbook.remove(external)
    sheets = workbook.find("s:sheets", NS)
    selected = list(sheets)[layout.sheet_index]
    selected_index = layout.sheet_index
    sheet_path = layout.sheet_path
    sheet = xml(parts[sheet_path])
    if sheet.find("s:oleObjects", NS) is not None:
        raise CertificateError(f"La hoja {layout.sheet_name} contiene objetos incrustados. Convierte esos objetos en imágenes en una copia del Excel.")
    headers = header_pictures(data, layout) if headers is None else headers
    take_format_footers(sheet)
    # Page-sized stationery is placed once, behind the content, after export.
    # Calc must not scale, crop or expand its header while printing the cells.
    header = sheet.find("s:headerFooter", NS)
    if header is not None:
        for picture in headers[0]:
            if picture.full_page:
                element = header.find(f"s:{picture.variant}Header", NS)
                if element is not None:
                    element.text = (element.text or "").replace("&G", "")
    # The removed page-sized header is restored behind the PDF later. If no
    # picture marker remains, its inactive VML and raster need not enter Calc.
    # Keep active header/footer pictures and their original text untouched.
    legacy = sheet.find("s:legacyDrawingHF", NS)
    active_pictures = header is not None and any("&G" in (element.text or "") for element in header)
    if headers[0] and legacy is not None and not active_pictures:
        rel_id = legacy.get(f"{{{R}}}id")
        sheet.remove(legacy)
        rel_path = relationship_file(sheet_path)
        relationships = xml(parts[rel_path])
        for relation in list(relationships):
            if relation.get("Id") == rel_id:
                relationships.remove(relation)
        parts[rel_path] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
    strings = []
    if "xl/sharedStrings.xml" in parts:
        strings = ["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
    cells = sheet.findall("s:sheetData/s:row/s:c", NS)
    queries = sheet.find("s:queryTableParts", NS)
    if queries is not None:
        sheet.remove(queries)
    guide = next((cell for cell in cells if "escanee este qr" in cell_value(cell, strings).casefold()), None)
    author = next((cell for cell in cells if "autorizado y firma" in cell_value(cell, strings).casefold()), None)
    printed = printed_cell_checker(workbook, sheet, selected_index, layout.sheet_name)
    omitted_auxiliary_refs = omit_auxiliary_errors(sheet, xml(parts["xl/styles.xml"]), printed)
    formula_errors = []
    for cell in cells:
        formula = cell.find("s:f", NS)
        if formula is not None:
            value = cell.find("s:v", NS)
            if value is None or (value.text is None and cell.get("t") != "str"):
                if printed(cell.get("r", "")):
                    raise CertificateError(f"Hay fórmulas sin resultados guardados en {layout.sheet_name}, celda {cell.get('r')}. Abre el archivo en Excel, calcula y guarda antes de subirlo.")
            cell.remove(formula)
        if cell.get("t") == "e" and printed(cell.get("r", "")):
            value = cell.find("s:v", NS)
            formula_errors.append(f"{cell.get('r')} ({value.text if value is not None else 'error'})")
    if formula_errors:
        raise CertificateError(f"Hay errores de fórmula en {layout.sheet_name}: {', '.join(formula_errors[:8])}. Corrige esas celdas y guarda el Excel antes de subirlo.")
    for row in sheet.findall("s:sheetData/s:row", NS):
        if row.get("ht"):
            row.set("customHeight", "1")
    heading_ranges = ()
    heading_drawing = sheet.find('s:drawing',NS)
    if heading_drawing is not None:
        heading_relations = xml(parts.get(relationship_file(sheet_path),b'<Relationships/>'))
        relation = next((rel for rel in heading_relations if rel.get('Id')==heading_drawing.get(f'{{{R}}}id')),None)
        heading_ranges = drawing_row_ranges(xml(parts[related(sheet_path,relation.get('Target'))]),sheet) if relation is not None else None
    for other in list(sheets):
        if other is not selected:
            sheets.remove(other)
    selected.set("state", "visible")
    view = workbook.find("s:bookViews/s:workbookView", NS)
    if view is not None:
        view.set("activeTab", "0")
        view.set("firstSheet", "0")
    names = workbook.find("s:definedNames", NS)
    if names is None:
        names = ET.SubElement(workbook, f"{{{S}}}definedNames")
    area_exists = False
    chosen_names = list(print_names(workbook, selected_index).values())
    for name in list(names):
        if name in chosen_names:
            if not name.get("name").startswith("_xlnm."):
                name.set("name", "_xlnm." + name.get("name"))
            name.set("localSheetId", "0")
            area_exists |= name.get("name") == "_xlnm.Print_Area"
        else:
            names.remove(name)
    if not area_exists:
        # Some saved templates (including CERTIFICADO QCP) have no Print_Area.
        # Include meaningful cells and merged ranges, not stray formatted columns.
        refs = [cell.get("r") for cell in cells if cell_value(cell, strings).strip()]
        refs += [merge.get("ref", "").split(":")[-1] for merge in sheet.findall("s:mergeCells/s:mergeCell", NS)]
        coords = [re.fullmatch(r"([A-Z]+)([0-9]+)", ref or "") for ref in refs]
        coords = [match for match in coords if match]
        if not coords:
            raise CertificateError("La hoja del certificado no contiene datos para imprimir.")
        def column_number(value):
            result = 0
            for character in value:
                result = result * 26 + ord(character) - 64
            return result
        last_col = max((match[1] for match in coords), key=column_number)
        last_row = max(int(match[2]) for match in coords)
        name = selected.get("name").replace("'", "''")
        ET.SubElement(names, f"{{{S}}}definedName", name="_xlnm.Print_Area", localSheetId="0").text = f"'{name}'!$A$1:${last_col}${last_row}"
        # An inferred range has no saved horizontal print placement. Centre it
        # within the existing margins instead of leaving all spare width right.
        options = sheet.find("s:printOptions", NS)
        if options is None:
            options = ET.Element(f"{{{S}}}printOptions")
            page_margins = sheet.find("s:pageMargins", NS)
            position = list(sheet).index(page_margins) if page_margins is not None else len(sheet)
            sheet.insert(position, options)
        options.set("horizontalCentered", "1")
    # The QCP zero-margin background is rendered as a single page image, not tiled.
    background = letterhead.data if letterhead else None
    footer = ""
    sheet_rels_path = relationship_file(sheet_path)
    sheet_rels = xml(parts[sheet_rels_path]) if sheet_rels_path in parts else ET.Element("rels")
    picture = sheet.find("s:picture", NS)
    if picture is not None:
        picture_rel = next((rel for rel in sheet_rels if rel.get("Id") == picture.get(f"{{{R}}}id")), None)
        if picture_rel is not None and letterhead is None:
            background = parts[related(sheet_path, picture_rel.get("Target"))]
            with Image.open(BytesIO(background)) as image:
                if image.width * image.height > 12_000_000:
                    raise CertificateError("El fondo del Excel es demasiado grande.")
        if picture_rel is not None and letterhead:
            sheet_rels.remove(picture_rel)
            parts[sheet_rels_path] = ET.tostring(sheet_rels, encoding="utf-8", xml_declaration=True)
        sheet.remove(picture)
    footer_element = sheet.find("s:headerFooter/s:oddFooter", NS)
    # The original zero-margin QCP background needs the existing footer offset.
    # Other headers/footers keep their native styles, alignment and page variants.
    margins = sheet.find("s:pageMargins", NS)
    zero_margin_background = background and margins is not None and all(
        float(margins.get(k, "0")) == 0 for k in ("left", "right", "top", "bottom"))
    if footer_element is not None and zero_margin_background and "&G" not in (footer_element.text or ""):
        footer = footer_element.text or ""
        footer_element.text = ""
    options = sheet.find("s:printOptions", NS)
    if options is not None and zero_margin_background:
        options.set("horizontalCentered", "0")
        options.set("verticalCentered", "0")
    missing_margins = margins is None
    if letterhead and missing_margins:
        margins = ET.SubElement(sheet, f"{{{S}}}pageMargins")
    if (letterhead and missing_margins) or zero_margin_background:
        # Leave room for the full-page QCP letterhead and footer from the example.
        margins.attrib.update(left="0.5", right="0.5", top="0.4", bottom="0.18", header="0", footer="0.18")
    setup = sheet.find("s:pageSetup", NS)
    if letterhead and setup is not None and (setup.get("paperSize", "9") != "9" or setup.get("orientation", "portrait") != "portrait"):
        raise CertificateError("Los fondos QCP requieren páginas A4 verticales. Revisa el tamaño y la orientación del Excel.")
    if letterhead and setup is None:
        setup = ET.SubElement(sheet, f"{{{S}}}pageSetup", paperSize="9", orientation="portrait")
    if letterhead and setup is not None:
        setup.set("paperSize", "9")
        setup.set("orientation", "portrait")
    if setup is not None:
        # Reserving side margins for a worksheet background can otherwise push
        # the rightmost columns onto separate pages. Keep fixed-scale layouts
        # elsewhere: fitting them can enlarge cells and add vertical pages.
        overflowing_background = zero_margin_background and background_width_overflows(workbook, sheet, layout.sheet_name)
        if overflowing_background:
            margins.set("bottom", "0.62")
        if overflowing_background or int(setup.get("fitToWidth", "1")) > 1:
            setup.set("fitToWidth", "1")
            setup.set("fitToHeight", "0")
            properties = sheet.find("s:sheetPr", NS)
            if properties is None:
                properties = ET.Element(f"{{{S}}}sheetPr")
                sheet.insert(0, properties)
            page_properties = properties.find("s:pageSetUpPr", NS)
            if page_properties is None:
                page_properties = ET.SubElement(properties, f"{{{S}}}pageSetUpPr")
            page_properties.set("fitToPage", "1")
        for key in list(setup.attrib):
            if key.startswith("{"):
                del setup.attrib[key]
    styles = xml(parts['xl/styles.xml'])
    fitted_overflow = fit_overflow_labels(sheet,styles,strings,cell_value,printed,drawing_ranges=heading_ranges)
    wrap_table_lines(sheet,styles,strings,cell_value,printed)
    fitted_fonts = fit_table_fonts(sheet,styles,strings,cell_value,printed)
    fitted_headings = fit_table_headings(sheet,styles,strings,cell_value,printed,drawing_ranges=heading_ranges)
    parts['xl/styles.xml'] = ET.tostring(styles,encoding='utf-8',xml_declaration=True)
    warnings = []
    if fitted_overflow:
        warnings.append('Se ajustaron rótulos largos dentro de sus celdas vacías contiguas, sin mover datos ni filas.')
    if fitted_fonts:
        warnings.append('Se ajustó el tamaño de letra de las tablas para conservar su legibilidad con la escala de impresión guardada.')
    if fitted_headings:
        warnings.append("Se ajustó la altura de los textos de tabla usando espacios vacíos cercanos, sin cambiar la paginación: " + ", ".join(fitted_headings) + ".")
    if layout.selection_reason:
        warnings.append(layout.selection_reason)
    if omitted_auxiliary_refs:
        warnings.append(f"Se omitieron errores auxiliares en celdas con texto blanco y fórmula oculta de {layout.sheet_name}: {', '.join(omitted_auxiliary_refs)}. El procesamiento continuó conservando los datos de medición.")
    if any("vbaproject" in name.lower() for name in parts):
        warnings.append("Se imprimieron los valores guardados del Excel sin ejecutar sus macros. El archivo original no se modificó.")
    embedded_signature = False
    drawing_ranges = ()
    drawing = sheet.find("s:drawing", NS)
    if drawing is not None:
        drawing_rel = next((rel for rel in sheet_rels if rel.get("Id") == drawing.get(f"{{{R}}}id")), None)
        if drawing_rel is not None:
            drawing_path = related(sheet_path, drawing_rel.get("Target"))
            drawing_root = xml(parts[drawing_path])
            image_rels = xml(parts.get(relationship_file(drawing_path), b"<Relationships/>"))
            images = {rel.get("Id"): related(drawing_path, rel.get("Target")) for rel in image_rels}
            signature_guide = guide if guide is not None else author
            guide_row = int(re.search(r"\d+", signature_guide.get("r")).group()) - 1 if signature_guide is not None else None
            candidates = []
            for anchor in drawing_root:
                point = anchor.find("d:from", NS)
                blip = anchor.find(".//a:blip", NS)
                if point is not None and blip is not None:
                    row = int(point.find("d:row", NS).text)
                    col = int(point.find("d:col", NS).text)
                    candidates.append((anchor, row, col, images.get(blip.get(f"{{{R}}}embed"))))
            near_signature = [candidate for candidate in candidates if guide_row is not None and abs(candidate[1] - guide_row) <= 2]
            signature_anchor = max(near_signature, key=lambda entry: entry[2], default=None) if author is not None else None
            removed_seals = 0
            for anchor, row, col, image_path in candidates:
                if signature_anchor is not None and anchor is signature_anchor[0]:
                    if replace_signature:
                        drawing_root.remove(anchor)
                    else:
                        embedded_signature = True
                    continue
                if guide_row is None or row < guide_row - 2 or not image_path or not image_path.lower().endswith((".png", ".jpeg", ".jpg")):
                    continue
                with Image.open(BytesIO(parts[image_path])) as image:
                    if 0.85 <= image.width / image.height <= 1.15:
                        drawing_root.remove(anchor)
                        removed_seals += 1
            if removed_seals:
                warnings.append("Se omitieron los sellos del Excel para conservar solo la firma, como en la referencia.")
            parts[drawing_path] = ET.tostring(drawing_root, encoding="utf-8", xml_declaration=True)
            drawing_ranges = drawing_row_ranges(drawing_root, sheet)
    composed = compose_numbered_pages(workbook, sheet, layout.sheet_name,
                                     xml(parts["xl/styles.xml"]), strings, cell_value, drawing_ranges)
    if composed:
        warnings.append(f"Se organizaron {len(composed)} páginas según la numeración de la plantilla, conservando los cuadros completos y reduciendo espacios vacíos excesivos.")
    parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
    parts[sheet_path] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
    return print_package(parts, selected.get(f"{{{R}}}id")), background, footer, embedded_signature, warnings


def convert_excel(data: bytes, replace_signature: bool = False,
                  letterhead: Letterhead | None = None, format_code: str = '', format_version: str = '') -> ConvertedExcel:
    from app.excel_sheets import certificate_layouts
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise CertificateError("La conversión de Excel necesita LibreOffice instalado en el servidor.")
    try:
        layouts = certificate_layouts(data)
        records, warnings = [], []
        original_parts = _read_parts(data)
        fallback_footer = general_data_footer(original_parts) or entered_footer(format_code, format_version)
        primary_footer = take_format_footers(xml(original_parts[layouts[0].sheet_path]))
        fallback_footer = primary_footer.get('odd') or next(iter(primary_footer.values()),None) or fallback_footer
        for index, layout in enumerate(layouts):
            selected = None if layout.preserve_format else letterhead
            headers = header_pictures(data, layout)
            workbook, background, footer, embedded, notes = _prepare_workbook(
                data, replace_signature and index == 0, selected, headers, layout)
            settings = take_format_footers(xml(original_parts[layout.sheet_path]))
            if fallback_footer:
                settings.setdefault('odd',fallback_footer)
            records.append((layout, workbook, background, footer, embedded, headers, selected, settings))
            warnings.extend(notes)
        if len(layouts) > 1:
            warnings.append(f"Se incluyeron las {len(layouts)} hojas del certificado en el orden de su numeración.")
        if any(not record[7] for record in records):
            warnings.append('No se encontró el código y la versión del formato del certificado en el pie ni en Datos generales. Se conservó la numeración; puedes completar esos datos en Pie del formato sin cambiar las mediciones.')
        numbered_total = max((n for record in records for n in record[0].page_numbers), default=0)
        del original_parts
    except (KeyError, IndexError, ValueError, OSError, AttributeError, TypeError, zipfile.BadZipFile) as exc:
        if isinstance(exc, CertificateError):
            raise
        raise CertificateError("No se pudo leer la estructura de este Excel.") from exc
    with tempfile.TemporaryDirectory(prefix="qcp-excel-") as directory:
        work = Path(directory)
        inputs = []
        for index, record in enumerate(records):
            path = work / f"certificate-{index}.xlsx"
            path.write_bytes(record[1])
            inputs.append(str(path))
        # Drop decoded preview rasters before starting the Office subprocess.
        # Keep the compressed, original-resolution images in the print files.
        pymupdf.TOOLS.store_shrink(100)
        deadline = monotonic() + 120
        def export(paths):
            remaining = deadline - monotonic()
            try:
                if remaining <= 0:
                    raise subprocess.TimeoutExpired(binary, 120)
                result = subprocess.run([
                    binary, f"-env:UserInstallation={(work / 'profile').as_uri()}",
                    "--headless", "--nologo", "--nodefault", "--nofirststartwizard", "--norestore",
                    "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), *paths,
                ], capture_output=True, timeout=remaining, env={**os.environ, "XDG_CACHE_HOME": str(work / "cache")})
            except subprocess.TimeoutExpired as exc:
                raise CertificateError("La conversión del Excel tardó demasiado. Prueba con un libro más pequeño.") from exc
            if result.returncode or any(not Path(path).with_suffix(".pdf").is_file() for path in paths):
                raise CertificateError("LibreOffice no pudo convertir este Excel. Revisa su formato y área de impresión.")
        export(inputs)
        repaired, retry = set(), []
        for index, record in enumerate(records):
            layout = record[0]
            with pymupdf.open(Path(inputs[index]).with_suffix('.pdf')) as section:
                mismatch = layout.page_numbers and len(section) != len(layout.page_numbers)
            if mismatch:
                workbook = fit_manual_numbered_pages(record[1], layout)
                if workbook is not None:
                    Path(inputs[index]).write_bytes(workbook)
                    Path(inputs[index]).with_suffix('.pdf').unlink()
                    retry.append(inputs[index])
                    repaired.add(index)
        if retry:
            export(retry)
            warnings.append('Se ajustó la impresión a los bloques de páginas definidos por los saltos guardados, conservando sus datos, firma e imágenes.')
        with pymupdf.open() as doc:
            background_digests, letterhead_pages = [], []
            format_footers = []
            for sheet_index, record in enumerate(records):
                layout, _, background, footer, _, headers, selected, footer_settings = record
                with pymupdf.open(Path(inputs[sheet_index]).with_suffix(".pdf")) as section:
                    if len(section) < 1 or len(doc) + len(section) > MAX_PAGES:
                        raise CertificateError(f"El certificado convertido debe tener entre 1 y {MAX_PAGES} páginas.")
                    if layout.page_numbers and len(section) != len(layout.page_numbers):
                        raise CertificateError(f"La hoja {layout.sheet_name} se dividió en {len(section)} páginas al imprimir. Revisa su área, escala y saltos de impresión para conservar la numeración del certificado.")
                    fit_stationery_content(section, headers, added_letterhead=bool(selected or background))
                    if layout.page_numbers:
                        validate_page_labels(section, layout.page_numbers, numbered_total)
                    validate_table_text(section)
                    if sheet_index in repaired:
                        validate_saved_labels(section, Path(inputs[sheet_index]).read_bytes(), layout)
                    background_xref = 0
                    for index, page in enumerate(section):
                        if background and selected is None:
                            background_xref = page.insert_image(page.rect, stream=background if not background_xref else None,
                                                                xref=background_xref, keep_proportion=False, overlay=False)
                            background_digests.append(next(image["digest"] for image in page.get_image_info(hashes=True, xrefs=True)
                                                           if image["xref"] == background_xref))
                        if footer:
                            text = footer.replace("&P", str(index + 1)).replace("&N", str(len(section)))
                            text = re.sub(r'&[LCR]|&"[^"]*"|&\d+', "", text)
                            for line_index, line in enumerate(text.strip().splitlines()):
                                width = pymupdf.get_text_length(line, fontname="helv", fontsize=6.5)
                                page.insert_text((page.rect.width - width - 8, page.rect.height - 45 + line_index * 10), line, fontsize=6.5)
                    if selected:
                        apply_letterhead(section, selected)
                        background_digests.append(selected.digest)
                        letterhead_pages.extend(range(len(doc) + 1, len(doc) + len(section) + 1))
                    restore_header_pictures(section, headers)
                    format_footers.extend(footer_for_page(footer_settings, index, headers) for index in range(len(section)))
                    background_digests.extend(picture.digest for picture in headers[0])
                    doc.insert_pdf(section)
                    pymupdf.TOOLS.store_shrink(100)
            background_digests = tuple(dict.fromkeys(background_digests))
            apply_format_footers(doc, format_footers)
            compact_backgrounds(doc, background_digests)
            for page in doc:
                page.clean_contents()
            doc.subset_fonts()
            return ConvertedExcel(doc.tobytes(garbage=4, deflate=True, use_objstms=1), records[0][4],
                                  list(dict.fromkeys(warnings)), background_digests, tuple(letterhead_pages), tuple(format_footers))
