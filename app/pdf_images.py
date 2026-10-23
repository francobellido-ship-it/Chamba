"""Remove duplicated JPEG colour metadata while keeping PDF colours and samples."""

import re

import pymupdf


def _without_jpeg_icc(data: bytes) -> bytes:
    if not data.startswith(b"\xff\xd8"):
        return data
    result = bytearray(data[:2])
    position, removed = 2, False
    while position + 4 <= len(data):
        if data[position] != 255:
            return data
        marker = data[position + 1]
        if marker in {0xDA, 0xD9}:  # The compressed image scan is copied verbatim.
            return bytes(result) + data[position:] if removed else data
        length = int.from_bytes(data[position + 2:position + 4], "big")
        end = position + 2 + length
        if length < 2 or end > len(data):
            return data
        if marker == 0xE2 and data[position + 4:end].startswith(b"ICC_PROFILE\0"):
            removed = True
        else:
            result.extend(data[position:end])
        position = end
    return data


def compact_backgrounds(doc: pymupdf.Document, digests: tuple[bytes, ...]) -> None:
    if not digests:
        return
    # Resource references identify shared images without decoding every raster
    # on every page. Only ICC JPEG candidates need their pixel digest checked.
    xrefs = {image[0] for page in doc for image in page.get_images()}
    for xref in sorted(xrefs):
        if not xref or doc.xref_get_key(xref, "Filter") != ("name", "/DCTDecode"):
            continue
        kind, value = doc.xref_get_key(xref, "ColorSpace")
        if kind == "xref":
            value = doc.xref_object(int(value.split()[0]))
        # PDF's declared ICCBased colour space controls interpretation of
        # the samples. Keep that profile, its Decode array, and JPEG data.
        if not re.search(r"/ICCBased\s+\d+\s+0\s+R", value):
            continue
        pixmap = pymupdf.Pixmap(doc, xref)
        digest = pixmap.digest
        del pixmap
        pymupdf.TOOLS.store_shrink(100)
        if digest not in digests:
            continue
        original = doc.xref_stream_raw(xref)
        compact = _without_jpeg_icc(original)
        if compact != original:
            doc.update_stream(xref, compact, compress=False)
            doc.xref_set_key(xref, "Filter", "/DCTDecode")


def clear_image_paint(doc, xref):
    """Replace an image's paint with an empty form, leaving text paths intact.

    Unlike delete_image's transparent placeholder, this has no image rectangle
    that could be mistaken for an obstacle under a replacement QR/background.
    Callers must ensure every use of this image is intended to be removed.
    """
    doc.update_object(xref,'<< /Type /XObject /Subtype /Form /BBox [0 0 1 1] /Resources << >> >>')
    doc.update_stream(xref,b'')
