"""Align the visible QR square with its printed instruction block."""

from io import BytesIO
import re

from PIL import Image
import pymupdf
import zxingcpp

from app.pdf import CertificateError


def insert_automatic_qr(page, qr, url, background_digests):
    """Use measured whitespace for a numbered legacy template with no QR guide.

    Keep the first page, all its content and original artwork. Only a known
    full-page background is exempt from collision checks; seals are occupied.
    """
    if page.rotation:
        raise CertificateError('No se pudo ubicar el QR en la plantilla girada.')
    text = [pymupdf.Rect(word[:4]) for word in page.get_text('words')]
    images = page.get_image_info(hashes=True)
    if not text or not any(info['digest'] in background_digests
                           and pymupdf.Rect(info['bbox']).get_area() >= page.rect.get_area() * .9 for info in images):
        raise CertificateError('No se pudo identificar un espacio seguro para el QR en esta plantilla.')
    occupied = text + [pymupdf.Rect(info['bbox']) for info in images
                       if not (info['digest'] in background_digests
                               and pymupdf.Rect(info['bbox']).get_area() >= page.rect.get_area() * .9)]
    occupied.extend(pymupdf.Rect(drawing['rect']) for drawing in page.get_drawings())
    occupied.extend(pymupdf.Rect(link['from']) for link in page.get_links())
    occupied.extend(annotation.rect for annotation in (page.annots() or []))
    bottom = page.rect.height - 44
    body = [area for area in text if 102 <= area.y0 and area.y1 < bottom - 8]
    if not body:
        raise CertificateError('No se pudo identificar el cuerpo del certificado para ubicar el QR.')
    left = max(36, min(area.x0 for area in body))
    right = min(page.rect.width - 36, max(area.x1 for area in body))
    for y in range(int(bottom - 80), int(page.rect.height / 2), -6):
        for x in (right - 66, left, (left + right - 66) / 2):
            frame = pymupdf.Rect(x, y, x + 66, y + 80)
            padded = frame + (-4, -4, 4, 4)
            if not page.rect.contains(padded) or any(padded.intersects(area) for area in occupied):
                continue
            area = pymupdf.Rect(x, y, x + 66, y + 66)
            page.insert_image(area, stream=qr)
            page.insert_link({'kind': pymupdf.LINK_URI, 'from': area, 'uri': url})
            page.insert_text((x, y + 76), 'Consultar certificado', fontsize=6.5)
            return area
    raise CertificateError('No hay espacio libre para añadir el QR sin cubrir información. No se cambió la paginación ni se emitió un certificado incompleto.')


def instruction_block(page, guide):
    block = pymupdf.Rect(guide)
    lines = [(pymupdf.Rect(line["bbox"]), "".join(span["text"] for span in line["spans"]).strip())
             for item in page.get_text("dict", flags=pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_PRESERVE_IMAGES)["blocks"] for line in item.get("lines", [])]
    continuation = re.compile(r"^(?:para\s+(?:descarg|consult|verific)|(?:el\s+)?(?:certificado|documento)|con\s+(?:la\s+)?firma|electr[oó]nica|original\b|en\s+pdf|pdf\b)", re.I)
    for area, text in sorted(lines, key=lambda item: item[0].y0):
        if abs(area.x0 - guide.x0) > 4 or area.y0 < guide.y1 - 1:
            continue
        if area.y0 - block.y1 > max(18, area.height * 1.5) or area.y1 - guide.y0 > 100:
            break
        if not continuation.match(text):
            break
        block |= area
    return block


def qr_position(page, guide, qr):
    block = instruction_block(page, guide)
    with Image.open(BytesIO(qr)) as image:
        left, top, right, bottom = image.getchannel("A").getbbox()
        visible_height = 66 * (bottom - top) / image.height
        # Short captions align at their top; multi-line instructions are
        # centered against the black/white square, excluding its quiet zone.
        y = (block.y0 - 66 * top / image.height if block.height < visible_height * .65
             else (block.y0 + block.y1) / 2 - 33)
    return pymupdf.Rect(guide.x0 - 72, y, guide.x0 - 6, y + 66)


def verify_template_qr(doc, area, url, qr, guide=None):
    page = doc[0]
    if not page.rect.contains(area) or abs(area.width - area.height) > .01:
        raise CertificateError("El QR no quedó cuadrado y dentro de la primera página.")
    if guide is not None and max(abs(a-b) for a,b in zip(area, qr_position(page, guide, qr))) > .1:
        raise CertificateError("El QR no quedó alineado con su bloque de instrucciones.")
    images = page.get_textpage(flags=pymupdf.TEXT_PRESERVE_IMAGES).extractIMGINFO()
    if not any(max(abs(a-b) for a,b in zip(info['bbox'],area)) < .1 for info in images):
        raise CertificateError("La imagen del QR no coincide con el espacio de consulta verificado.")
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(3, 3), clip=area)
    decoded = zxingcpp.read_barcodes(Image.open(BytesIO(pixmap.tobytes("png"))), formats=zxingcpp.BarcodeFormat.QRCode)
    if [item.text for item in decoded] != [url]:
        raise CertificateError("No se pudo comprobar la lectura del QR del PDF final. Revisa el espacio reservado en la plantilla.")
    links = [(index, pymupdf.Rect(link["from"])) for index, page in enumerate(doc)
             for link in page.get_links() if link.get("uri") == url]
    if len(links) != 1 or links[0][0] != 0 or max(abs(a-b) for a,b in zip(links[0][1], area)) > .1:
        raise CertificateError("El QR debe aparecer una sola vez y enlazar desde la primera página.")
