"""Recover auxiliary formula errors without suppressing printed measurements."""

import re

from app.excel_ranges import _bounds, column_number

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def omit_auxiliary_errors(sheet, styles, printed) -> list[str]:
    cells = sheet.findall("s:sheetData/s:row/s:c", NS)
    formats = styles.find("s:cellXfs", NS)
    fonts = styles.find("s:fonts", NS)
    fills = styles.find("s:fills", NS)
    borders = styles.find("s:borders", NS)
    conditional = [_bounds(rule.get("sqref", "").replace(" ", ","), "")
                   for rule in sheet.findall("s:conditionalFormatting", NS)]
    omitted = []
    for cell in cells:
        ref = cell.get("r", "")
        value = cell.find("s:v", NS)
        formula = cell.find("s:f", NS)
        missing = formula is not None and (value is None or (value.text is None and cell.get("t") != "str"))
        if (cell.get("t") != "e" and not missing) or not printed(ref):
            continue
        coord = re.fullmatch(r"([A-Z]+)([0-9]+)", ref)
        if coord is None:
            continue
        col, row = column_number(coord[1]), int(coord[2])
        if any(ranges is None or any(left <= col <= right and top <= row <= bottom
                                     for left, top, right, bottom in ranges) for ranges in conditional):
            continue
        style = formats[int(cell.get("s", "0"))]
        # Only explicitly concealed, unpainted General-format helpers qualify.
        # Numeric/date formats, tables and conditional or coloured cells retain
        # the normal validation; their missing measurements cannot be invented.
        if style.get("numFmtId", "0") != "0":
            continue
        font = fonts[int(style.get("fontId", "0"))]
        color = font.find("s:color", NS)
        if color is None or color.get("rgb", "").upper() not in {"FFFFFF", "FFFFFFFF"}:
            continue
        protection = style.find("s:protection", NS)
        if protection is None or protection.get("hidden") not in {"1", "true"}:
            continue
        fill = fills[int(style.get("fillId", "0"))]
        pattern = fill.find("s:patternFill", NS)
        border = borders[int(style.get("borderId", "0"))]
        if fill.find("s:gradientFill", NS) is not None or (pattern is not None
                and pattern.get("patternType", "none") != "none") or any(side.get("style") for side in border):
            continue
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)
        omitted.append(ref)
    return omitted
