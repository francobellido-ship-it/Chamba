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
    seen = set()
    for page in doc:
        for image in page.get_image_info(hashes=True, xrefs=True):
            xref = image["xref"]
            if not xref or xref in seen or image["digest"] not in digests:
                continue
            seen.add(xref)
            if doc.xref_get_key(xref, "Filter") != ("name", "/DCTDecode"):
                continue
            kind, value = doc.xref_get_key(xref, "ColorSpace")
            if kind == "xref":
                value = doc.xref_object(int(value.split()[0]))
            # PDF's declared ICCBased colour space controls interpretation of
            # the samples. Keep that profile, its Decode array, and JPEG data.
            if not re.search(r"/ICCBased\s+\d+\s+0\s+R", value):
                continue
            original = doc.xref_stream_raw(xref)
            compact = _without_jpeg_icc(original)
            if compact != original:
                doc.update_stream(xref, compact, compress=False)
                doc.xref_set_key(xref, "Filter", "/DCTDecode")
