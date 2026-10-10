"""Export the saved values and print layout of a QCP XLSX certificate."""

from dataclasses import dataclass
from io import BytesIO
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image
import pymupdf

from app.pdf import CertificateError, MAX_PAGES
from app.backgrounds import Letterhead, apply_letterhead
from app.excel_headers import header_pictures, restore_header_pictures

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


@dataclass(frozen=True)
class ExcelLayout:
    sheet_name: str
    sheet_path: str
    sheet_index: int
    preserve_format: bool


def _read_parts(data: bytes) -> dict[str, bytes]:
    try:
        source = zipfile.ZipFile(BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise CertificateError("El archivo debe ser un Excel .xlsx válido y sin contraseña.") from exc
    with source:
        entries = source.infolist()
        if len(entries) > 2000 or sum(e.file_size for e in entries) > 80 * 1024 * 1024:
            raise CertificateError("El contenido del Excel supera el tamaño permitido.")
        if len({e.filename for e in entries}) != len(entries):
            raise CertificateError("El Excel contiene partes duplicadas.")
        if any(e.filename.startswith("/") or ".." in PurePosixPath(e.filename).parts for e in entries):
            raise CertificateError("El Excel contiene rutas internas inválidas.")
        if any(e.flag_bits & 1 or "vbaproject" in e.filename.lower() for e in entries):
            raise CertificateError("Solo se admiten archivos .xlsx sin macros ni contraseña.")
        parts = {entry.filename: source.read(entry) for entry in entries}
    if "xl/workbook.xml" not in parts:
        raise CertificateError("El archivo no contiene un libro Excel válido.")
    return parts


def _layout(parts: dict[str, bytes]) -> ExcelLayout:
    workbook = xml(parts["xl/workbook.xml"])
    rels = xml(parts["xl/_rels/workbook.xml.rels"])
    targets = {rel.get("Id"): related("xl/workbook.xml", rel.get("Target", ""))
               for rel in rels if rel.get("TargetMode") != "External"}
    strings = (["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
               if "xl/sharedStrings.xml" in parts else [])
    candidates = []
    for index, entry in enumerate(workbook.findall("s:sheets/s:sheet", NS)):
        name = entry.get("name", "")
        if not re.fullmatch(r"certificado(?:\s+qcp)?", name.strip(), re.I):
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
    marked = [candidate for candidate in candidates if candidate[0]]
    if len(marked) > 1 or (not marked and len(candidates) > 1):
        raise CertificateError("Hay varias hojas de certificado. Deja una sola hoja de certificado para identificar el formato correcto.")
    _, index, name, path, sheet = marked[0] if marked else candidates[0]
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
    return ExcelLayout(name, path, index, bool(existing))


def inspect_excel(data: bytes) -> ExcelLayout:
    try:
        return _layout(_read_parts(data))
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
        return "".join(cell.find("s:is", NS).itertext())
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
                      letterhead: Letterhead | None = None) -> tuple[bytes, bytes | None, str, bool, list[str]]:
    parts = _read_parts(data)
    layout = _layout(parts)
    if layout.preserve_format:
        letterhead = None
    # Export saved results offline. Data connections are not needed for printing.
    discarded = {name for name in parts if name.startswith(("xl/externalLinks/", "xl/queryTables/"))
                 or name == "xl/connections.xml"}
    if any(name.startswith("xl/embeddings/") for name in parts):
        raise CertificateError("El Excel contiene objetos incrustados. Guarda una copia sin esos elementos.")
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
    strings = []
    if "xl/sharedStrings.xml" in parts:
        strings = ["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
    cells = sheet.findall("s:sheetData/s:row/s:c", NS)
    queries = sheet.find("s:queryTableParts", NS)
    if queries is not None:
        sheet.remove(queries)
    guide = next((cell for cell in cells if "escanee este qr" in cell_value(cell, strings).casefold()), None)
    author = next((cell for cell in cells if "autorizado y firma" in cell_value(cell, strings).casefold()), None)
    for cell in cells:
        formula = cell.find("s:f", NS)
        if formula is not None:
            value = cell.find("s:v", NS)
            if value is None or (value.text is None and cell.get("t") != "str"):
                raise CertificateError("Hay fórmulas sin resultados guardados en Certificado. Abre el Excel en Excel, calcula y guarda antes de subirlo.")
            if cell.get("t") == "e":
                raise CertificateError("Hay errores de fórmula en Certificado. Corrígelos y guarda el Excel antes de subirlo.")
            cell.remove(formula)
    for row in sheet.findall("s:sheetData/s:row", NS):
        if row.get("ht"):
            row.set("customHeight", "1")
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
    for name in list(names):
        if name.get("localSheetId") == str(selected_index) and name.get("name") in {"_xlnm.Print_Area", "_xlnm.Print_Titles"}:
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
    if letterhead:
        # A new letterhead needs room for its printed logo and contact footer.
        # Existing stationery takes the other path and keeps its saved margins.
        margins.set("top", str(max(float(margins.get("top", "0")), 1.4)))
        margins.set("bottom", str(max(float(margins.get("bottom", "0")), .65)))
    setup = sheet.find("s:pageSetup", NS)
    if letterhead and setup is not None and (setup.get("paperSize", "9") != "9" or setup.get("orientation", "portrait") != "portrait"):
        raise CertificateError("Los fondos QCP requieren páginas A4 verticales. Revisa el tamaño y la orientación del Excel.")
    if letterhead and setup is None:
        setup = ET.SubElement(sheet, f"{{{S}}}pageSetup", paperSize="9", orientation="portrait")
    if letterhead and setup is not None:
        setup.set("paperSize", "9")
        setup.set("orientation", "portrait")
    if setup is not None:
        for key in list(setup.attrib):
            if key.startswith("{"):
                del setup.attrib[key]
    warnings = []
    embedded_signature = False
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
    parts["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
    parts[sheet_path] = ET.tostring(sheet, encoding="utf-8", xml_declaration=True)
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return output.getvalue(), background, footer, embedded_signature, warnings


def convert_excel(data: bytes, replace_signature: bool = False,
                  letterhead: Letterhead | None = None) -> ConvertedExcel:
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise CertificateError("La conversión de Excel necesita LibreOffice instalado en el servidor.")
    try:
        if inspect_excel(data).preserve_format:
            letterhead = None
        workbook, background, footer, embedded, warnings = _prepare_workbook(data, replace_signature, letterhead)
        headers = header_pictures(workbook)
    except (KeyError, IndexError, ValueError, OSError, AttributeError, TypeError, zipfile.BadZipFile) as exc:
        if isinstance(exc, CertificateError):
            raise
        raise CertificateError("No se pudo leer la estructura de este Excel.") from exc
    with tempfile.TemporaryDirectory(prefix="qcp-excel-") as directory:
        work = Path(directory)
        (work / "certificate.xlsx").write_bytes(workbook)
        try:
            result = subprocess.run([
                binary, f"-env:UserInstallation={(work / 'profile').as_uri()}",
                "--headless", "--nologo", "--nodefault", "--nofirststartwizard", "--norestore",
                "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(work / "certificate.xlsx"),
            ], capture_output=True, timeout=120, env={**os.environ, "XDG_CACHE_HOME": str(work / "cache")})
        except subprocess.TimeoutExpired as exc:
            raise CertificateError("La conversión del Excel tardó demasiado. Prueba con un libro más pequeño.") from exc
        pdf_path = work / "certificate.pdf"
        if result.returncode or not pdf_path.is_file():
            raise CertificateError("LibreOffice no pudo convertir este Excel. Revisa su formato y área de impresión.")
        with pymupdf.open(pdf_path) as doc:
            if not 1 <= len(doc) <= MAX_PAGES:
                raise CertificateError(f"El certificado convertido debe tener entre 1 y {MAX_PAGES} páginas.")
            background_xref = 0
            background_digests = [letterhead.digest] if letterhead else []
            for index, page in enumerate(doc):
                if background and letterhead is None:
                    background_xref = page.insert_image(page.rect, stream=background if not background_xref else None,
                                                        xref=background_xref, overlay=False)
                    if not background_digests:
                        background_digests.append(next(image["digest"] for image in page.get_image_info(hashes=True, xrefs=True)
                                                       if image["xref"] == background_xref))
                if footer:
                    text = footer.replace("&P", str(index + 1)).replace("&N", str(len(doc)))
                    text = re.sub(r'&[LCR]|&"[^"]*"|&\d+', "", text)
                    lines = text.strip().splitlines()
                    for line_index, line in enumerate(lines):
                        width = pymupdf.get_text_length(line, fontname="helv", fontsize=6.5)
                        page.insert_text((page.rect.width - width - 8, page.rect.height - 45 + line_index * 10), line, fontsize=6.5)
            if letterhead:
                apply_letterhead(doc, letterhead)
            restore_header_pictures(doc, headers)
            background_digests.extend(picture.digest for picture in headers[0])
            return ConvertedExcel(doc.tobytes(garbage=4, deflate=True), embedded, warnings, tuple(background_digests))
