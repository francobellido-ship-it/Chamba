"""Keep printed cell content clear of a full-page certificate letterhead."""

import pymupdf

from app.pdf import CertificateError


def fit_stationery_content(doc: pymupdf.Document, headers: tuple) -> None:
    if not any(picture.full_page for picture in headers[0]):
        return
    # Native Excel backgrounds and page-sized header pictures are different:
    # these templates store the stationery in the header, outside the cells.
    for index, page in enumerate(doc):
        areas = [pymupdf.Rect(word[:4]) for word in page.get_text("words")]
        areas += [pymupdf.Rect(image["bbox"]) for image in page.get_image_info()]
        for drawing in page.get_drawings():
            if drawing.get("color") is not None or (drawing.get("fill") is not None
                    and any(channel < .98 for channel in drawing["fill"])):
                areas.append(drawing["rect"])
        if not areas:
            continue
        top = max(0, min(area.y0 for area in areas) - 2)
        bottom = min(page.rect.height, max(area.y1 for area in areas) + 2)
        if bottom <= top:
            continue
        source = pymupdf.Rect(0, top, page.rect.width, bottom)
        available = page.rect.height - 44 - 102
        scale = min(1, available / source.height)
        if scale < .8:
            raise CertificateError(f"El contenido de la página {index + 1} no cabe entre el encabezado y el pie. Revisa el área y los saltos de impresión del Excel.")
        y = max(102, top)
        y = min(y, page.rect.height - 44 - source.height * scale)
        target = pymupdf.Rect(0, y, page.rect.width * scale, y + source.height * scale)
        # Transform the original vector stream directly. Nested page forms
        # interact poorly with image caches in some mobile PDF engines.
        page.wrap_contents()
        content = b"\n".join(doc.xref_stream(xref) for xref in page.get_contents())
        translate = page.rect.height * (1 - scale) + top * scale - y
        clip_y = page.rect.height - target.y1
        prefix = (f"q\n0 {clip_y:.6f} {target.width:.6f} {target.height:.6f} re W n\n"
                  f"{scale:.8f} 0 0 {scale:.8f} 0 {translate:.6f} cm\n").encode()
        replacement = doc.get_new_xref()
        doc.update_object(replacement, "<<>>")
        doc.update_stream(replacement, prefix + content + b"\nQ\n")
        page.set_contents(replacement)
