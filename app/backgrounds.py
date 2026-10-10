"""Official QCP letterheads and checks required before publishing Excel PDFs."""

from dataclasses import dataclass
from functools import lru_cache
import hashlib
from io import BytesIO
from pathlib import Path

from PIL import Image
import pymupdf

from app.pdf import CertificateError


ASSETS = Path(__file__).parent / "assets"
TEMPLATES = {
    "acreditado": ("Acreditado", ASSETS / "fondo-acreditado.jpeg",
                   "7dd52c0f6fcc7935a2e34a39cc719eefe3c5661f1d1513ba419f3afc68404844"),
    "no_acreditado": ("No acreditado", ASSETS / "fondo-no-acreditado.jpeg",
                      "912944df2230a73ebc5e11178df555d47aad0e86e718c1b84b806a06c251f1e9"),
}
A4_WIDTH, A4_HEIGHT = 595.276, 841.890
POSITION_TOLERANCE = 1.0  # PDF points; accommodates image/page aspect rounding.


@dataclass(frozen=True)
class Letterhead:
    kind: str
    label: str
    data: bytes
    digest: bytes
    width: int
    height: int


def background_choices() -> list[dict]:
    return [{"value": kind, "label": label, "available": bool(checksum and path.is_file())}
            for kind, (label, path, checksum) in TEMPLATES.items()]


def require_background_type(kind: str) -> str:
    if kind not in TEMPLATES:
        raise CertificateError("Selecciona Acreditado o No acreditado para cada Excel.")
    _, path, checksum = TEMPLATES[kind]
    if not checksum or not path.is_file():
        raise CertificateError("El fondo oficial seleccionado todavía no está disponible. Contacta al administrador.")
    return kind


@lru_cache(maxsize=2)
def load_letterhead(kind: str) -> Letterhead:
    require_background_type(kind)
    label, path, checksum = TEMPLATES[kind]
    try:
        data = path.read_bytes()
        if hashlib.sha256(data).hexdigest() != checksum:
            raise CertificateError("El fondo oficial cambió o está dañado. No se emitirá el certificado.")
        with Image.open(BytesIO(data)) as image:
            if image.format not in {"PNG", "JPEG"} or not 1 < image.width * image.height <= 12_000_000:
                raise CertificateError("El archivo del fondo oficial no es una imagen válida.")
            width, height = image.size
            if abs(width / height - A4_WIDTH / A4_HEIGHT) > 0.002:
                raise CertificateError("El fondo oficial debe corresponder a una página A4 vertical.")
        # Fingerprint decoded PDF pixels, not reserialized JPEG headers.
        with pymupdf.open() as reference:
            page = reference.new_page(width=A4_WIDTH, height=A4_HEIGHT)
            page.insert_image(page.rect, stream=data, overlay=False)
            digest = page.get_image_info(hashes=True)[0]["digest"]
        return Letterhead(kind, label, data, digest, width, height)
    except CertificateError:
        raise
    except (OSError, ValueError, pymupdf.FileDataError) as exc:
        raise CertificateError("No se pudo leer el fondo oficial. No se emitirá el certificado.") from exc


def _check_page_size(page: pymupdf.Page) -> None:
    if (page.rotation or abs(page.rect.width - A4_WIDTH) > POSITION_TOLERANCE
            or abs(page.rect.height - A4_HEIGHT) > POSITION_TOLERANCE):
        raise CertificateError("Los fondos QCP requieren páginas A4 verticales. Revisa el tamaño y la orientación del Excel.")


def _opaque_page_image(info: dict, page: pymupdf.Page) -> bool:
    # delete_image leaves a tiny transparent replacement in old content streams.
    return (info["width"] > 2 and info["height"] > 2
            and pymupdf.Rect(info["bbox"]).get_area() > page.rect.get_area() * 0.9)


def apply_letterhead(doc: pymupdf.Document, selected: Letterhead) -> None:
    known = {load_letterhead(choice["value"]).digest for choice in background_choices() if choice["available"]}
    remove = {}
    for index, page in enumerate(doc):
        _check_page_size(page)
        for info in page.get_image_info(hashes=True, xrefs=True):
            if info["digest"] in known:
                if not info["xref"]:
                    raise CertificateError("El Excel contiene un fondo que no se puede reemplazar. Revisa la plantilla.")
                remove.setdefault(info["xref"], index)
            elif _opaque_page_image(info, page):
                raise CertificateError("El Excel contiene una imagen de página completa que puede tapar el fondo. Revisa la plantilla.")
    for xref, page_index in remove.items():
        doc[page_index].delete_image(xref)
    background_xref = 0
    for page in doc:
        background_xref = page.insert_image(page.rect, stream=selected.data if not background_xref else None,
                                            xref=background_xref, overlay=False)
    verify_letterhead(doc, selected)


def verify_letterhead(doc: pymupdf.Document, selected: Letterhead) -> list[int]:
    verified = []
    for index, page in enumerate(doc):
        _check_page_size(page)
        images = page.get_image_info(hashes=True, xrefs=True)
        matches = [info for info in images if info["digest"] == selected.digest]
        if len(matches) != 1:
            raise CertificateError(f"La página {index + 1} no tiene exactamente un fondo oficial. No se emitirá el certificado.")
        image = matches[0]
        if (image["width"], image["height"]) != (selected.width, selected.height):
            raise CertificateError("La resolución del fondo oficial cambió. No se emitirá el certificado.")
        area = pymupdf.Rect(image["bbox"])
        if any(abs(actual - expected) > POSITION_TOLERANCE for actual, expected in zip(area, page.rect)):
            raise CertificateError(f"El fondo de la página {index + 1} está desfasado o incompleto. No se emitirá el certificado.")
        if any(info["digest"] != selected.digest and _opaque_page_image(info, page) for info in images):
            raise CertificateError(f"Una imagen tapa el fondo de la página {index + 1}. Revisa la plantilla.")
        paint = page.get_bboxlog()
        if not paint or paint[0][0] != "fill-image" or any(
                abs(actual - expected) > POSITION_TOLERANCE for actual, expected in zip(paint[0][1], area)):
            raise CertificateError(f"El fondo de la página {index + 1} no está detrás del contenido. No se emitirá el certificado.")
        if any(drawing.get("fill") is not None and drawing.get("fill_opacity", 1) >= 0.99
               and drawing["rect"].get_area() > page.rect.get_area() * 0.9 for drawing in page.get_drawings()):
            raise CertificateError(f"El contenido de la página {index + 1} tapa el fondo completo. Revisa la plantilla.")
        verified.append(index + 1)
    return verified
