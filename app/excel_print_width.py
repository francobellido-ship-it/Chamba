"""Detect width overflow before enabling automatic print scaling."""

import math

from app.excel_ranges import NS, _bounds, print_names


def background_width_overflows(workbook, sheet, sheet_name: str) -> bool:
    setup = sheet.find("s:pageSetup", NS)
    margins = sheet.find("s:pageMargins", NS)
    area = print_names(workbook, 0).get("Print_Area")
    ranges = _bounds(area.text or "", sheet_name) if area is not None else None
    if setup is None or margins is None or not ranges or len(ranges) != 1:
        return False
    if setup.get("paperSize", "9") != "9" or setup.get("orientation", "portrait") != "portrait":
        return False
    first, _, last, _ = ranges[0]
    format_properties = sheet.find("s:sheetFormatPr", NS)
    default = float(format_properties.get("defaultColWidth", "8.43")) if format_properties is not None else 8.43
    widths = dict.fromkeys(range(first, last + 1), default)
    for column in sheet.findall("s:cols/s:col", NS):
        value = 0 if column.get("hidden") in {"1", "true"} else float(column.get("width", str(default)))
        for index in range(max(first, int(column.get("min"))), min(last, int(column.get("max"))) + 1):
            widths[index] = value
    # Standard Excel character widths include the cell's five-pixel padding
    # when rendered by Calc. Only shrink a fixed-scale background template
    # when the estimated width exceeds its newly reserved side margins.
    points = sum(math.floor(width * 7 + 5) * .75 for width in widths.values() if width > 0)
    available = 595.276 - 72 * (float(margins.get("left", "0")) + float(margins.get("right", "0")))
    return points * float(setup.get("scale", "100")) / 100 > available + 1
