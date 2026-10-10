"""Preserve VML header/footer pictures that Calc omits from XLSX exports."""

from dataclasses import dataclass
from io import BytesIO
import re

from PIL import Image
import pymupdf

from app.pdf import CertificateError

V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"


@dataclass(frozen=True)
class HeaderPicture:
    data: bytes
    section: str
    variant: str
    width: float
    height: float
    dx: float
    dy: float
    crops: tuple[float, float, float, float]
    digest: bytes


def _points(value: str) -> float:
    match = re.fullmatch(r"\s*(-?[0-9.]+)(pt|px|in|cm|mm)?\s*", value)
    if not match:
        raise CertificateError("No se pudo leer el tamaño de una imagen del encabezado o pie del Excel.")
    return float(match[1]) * {None: 1, "pt": 1, "px": .75, "in": 72, "cm": 72 / 2.54, "mm": 72 / 25.4}[match[2]]


def header_pictures(workbook: bytes) -> tuple[list[HeaderPicture], dict[str, float], bool, bool]:
    # Imported lazily because the converter calls this module after initialization.
    from app.excel import _read_parts, _layout, xml, related, relationship_file, NS, R
    parts = _read_parts(workbook)
    layout = _layout(parts)
    sheet = xml(parts[layout.sheet_path])
    header = sheet.find("s:headerFooter", NS)
    drawing = sheet.find("s:legacyDrawingHF", NS)
    margins = sheet.find("s:pageMargins", NS)
    positions = {key: float(margins.get(key, "0")) * 72 for key in ("left", "right", "header", "footer")} if margins is not None else dict(left=36, right=36, header=21.6, footer=21.6)
    if header is None or not any("&G" in (e.text or "") for e in header):
        return [], positions, False, False
    if drawing is None:
        raise CertificateError("El encabezado o pie referencia una imagen que falta en el Excel. Guarda una copia con la imagen incorporada.")
    rels = xml(parts[relationship_file(layout.sheet_path)])
    target = next((rel.get("Target") for rel in rels if rel.get("Id") == drawing.get(f"{{{R}}}id")), None)
    if not target:
        raise CertificateError("Falta la imagen del encabezado o pie dentro del Excel.")
    path = related(layout.sheet_path, target)
    root = xml(parts[path])
    images = {rel.get("Id"): related(path, rel.get("Target")) for rel in xml(parts[relationship_file(path)])}
    scale = 1.0
    setup = sheet.find("s:pageSetup", NS)
    if header.get("scaleWithDoc", "1") != "0" and setup is not None:
        scale = float(setup.get("scale", "100")) / 100
    pictures = []
    for shape in root.findall(f".//{{{V}}}shape"):
        identity = re.fullmatch(r"([LCR][HF])([EF]?)", shape.get("id", ""))
        if not identity:
            continue
        image = shape.find(f"{{{V}}}imagedata")
        if image is None:
            continue
        rel_id = image.get(f"{{{O}}}relid") or image.get(f"{{{R}}}id")
        data = parts[images[rel_id]]
        with Image.open(BytesIO(data)) as raster:
            if raster.width * raster.height > 12_000_000:
                raise CertificateError("La imagen del encabezado o pie es demasiado grande.")
        style = dict(item.strip().split(":", 1) for item in shape.get("style", "").split(";") if ":" in item)
        width, height = _points(style.get("width", "0")) * scale, _points(style.get("height", "0")) * scale
        if not 0 < width <= 2000 or not 0 < height <= 2000:
            raise CertificateError("El tamaño de una imagen del encabezado o pie no es válido.")
        crops = []
        for side in ("cropleft", "croptop", "cropright", "cropbottom"):
            value = image.get(side, "0")
            crop = float(value[:-1]) / 65536 if value.endswith("f") else float(value)
            if not 0 <= crop < 1:
                raise CertificateError("El recorte de una imagen del encabezado o pie no es válido.")
            crops.append(crop)
        if crops[0] + crops[2] >= 1 or crops[1] + crops[3] >= 1:
            raise CertificateError("El recorte del encabezado o pie oculta toda su imagen.")
        with pymupdf.open() as reference:
            reference.new_page().insert_image(pymupdf.Rect(0, 0, width, height), stream=data)
            digest = reference[0].get_image_info(hashes=True)[0]["digest"]
        pictures.append(HeaderPicture(data, identity[1], {"": "odd", "E": "even", "F": "first"}[identity[2]],
                                      width, height, _points(style.get("margin-left", "0")) * scale,
                                      _points(style.get("margin-top", "0")) * scale, tuple(crops), digest))
    if not pictures:
        raise CertificateError("No se pudo leer la imagen del encabezado o pie del Excel.")
    return pictures, positions, header.get("differentFirst", "0") == "1", header.get("differentOddEven", "0") == "1"


def restore_header_pictures(doc: pymupdf.Document, settings: tuple) -> None:
    pictures, margins, first, even = settings
    for index, page in enumerate(doc):
        variant = "first" if first and index == 0 else "even" if even and (index + 1) % 2 == 0 else "odd"
        # Read a fresh text page: get_image_info caches data across insertions.
        native = page.get_textpage(flags=pymupdf.TEXT_PRESERVE_IMAGES).extractIMGINFO(hashes=True)
        for picture in pictures:
            if picture.variant != variant:
                continue
            x = {"L": margins["left"], "C": (page.rect.width - picture.width) / 2,
                 "R": page.rect.width - margins["right"] - picture.width}[picture.section[0]] + picture.dx
            y = (margins["header"] if picture.section[1] == "H" else
                 page.rect.height - margins["footer"] - picture.height) + picture.dy
            area = pymupdf.Rect(x, y, x + picture.width, y + picture.height)
            if any(image["digest"] == picture.digest and all(abs(a - b) < .1 for a, b in zip(image["bbox"], area)) for image in native):
                continue
            if any(picture.crops):
                # PDF clipping retains the source raster and its original pixels.
                with pymupdf.open() as image_pdf:
                    image_pdf.new_page(width=picture.width, height=picture.height).insert_image(
                        pymupdf.Rect(0, 0, picture.width, picture.height), stream=picture.data, keep_proportion=False)
                    left, top, right, bottom = picture.crops
                    clip = pymupdf.Rect(left * picture.width, top * picture.height,
                                        (1 - right) * picture.width, (1 - bottom) * picture.height)
                    page.show_pdf_page(area, image_pdf, clip=clip, keep_proportion=False, overlay=False)
            else:
                page.insert_image(area, stream=picture.data, keep_proportion=False, overlay=False)
            images = page.get_textpage(flags=pymupdf.TEXT_PRESERVE_IMAGES).extractIMGINFO(hashes=True)
            if picture.digest not in {image["digest"] for image in images}:
                raise CertificateError("No se pudo conservar una imagen del encabezado o pie. No se emitirá el certificado.")
