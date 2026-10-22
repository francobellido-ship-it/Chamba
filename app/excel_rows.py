"""Make wrapped table headings readable without moving saved page boundaries."""

import math
import re

import pymupdf

from app.excel_ranges import NS


def fit_single_line_labels(sheet, styles):
    """Fit merged one-line labels without changing their data or row geometry."""
    from copy import deepcopy
    formats = styles.find('s:cellXfs', NS)
    borders = styles.findall('s:borders/s:border', NS)
    merged = {merge.get('ref').split(':')[0] for merge in sheet.findall('s:mergeCells/s:mergeCell', NS)
              if len({re.sub(r'[A-Z]+', '', ref) for ref in merge.get('ref').split(':')}) == 1}
    replacements = {}
    for cell in sheet.findall('s:sheetData/s:row/s:c', NS):
        if cell.get('r') not in merged or cell.get('t') not in {'s', 'str', 'inlineStr'}:
            continue
        index = int(cell.get('s', '0'))
        style = formats[index]
        alignment = style.find('s:alignment', NS)
        if (alignment is None or alignment.get('wrapText') in {'1', 'true'}
                or alignment.get('textRotation', '0') != '0' or alignment.get('shrinkToFit') in {'1', 'true'}
                or not any(edge.get('style') for edge in borders[int(style.get('borderId', '0'))])):
            continue
        if index not in replacements:
            copy = deepcopy(style)
            copy.find('s:alignment', NS).set('shrinkToFit', '1')
            replacements[index] = len(formats)
            formats.append(copy)
        cell.set('s', str(replacements[index]))
    formats.set('count', str(len(formats)))


def _coordinate(reference):
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", reference)
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    return column, int(match[2])


def fit_table_headings(sheet, styles, strings, value_of, printed):
    """Borrow only empty preceding spacers; never grow the printed page.

    Calc does not auto-fit horizontally merged cells. Use conservative Arial
    metrics for the common Arial/Calibri headings, leaving fonts and data intact.
    Unsupported rotations, vertical merges and drawings retain their layout.
    """
    if sheet.find("s:drawing", NS) is not None:
        return []
    fonts = styles.findall("s:fonts/s:font", NS)
    formats = styles.findall("s:cellXfs/s:xf", NS)
    fills = styles.findall("s:fills/s:fill", NS)
    borders = styles.findall("s:borders/s:border", NS)
    properties = sheet.find("s:sheetFormatPr", NS)
    default_height = float(properties.get("defaultRowHeight", "15")) if properties is not None else 15
    default_width = float(properties.get("defaultColWidth", "8.43")) if properties is not None else 8.43
    columns = sheet.findall("s:cols/s:col", NS)
    merges = []
    for merge in sheet.findall("s:mergeCells/s:mergeCell", NS):
        references = merge.get("ref").split(":")
        first, last = references[0], references[-1]
        merges.append((*_coordinate(first), *_coordinate(last)))
    rows = {int(row.get("r")): row for row in sheet.findall("s:sheetData/s:row", NS)}
    breaks = {int(brk.get("id")) for brk in sheet.findall("s:rowBreaks/s:brk", NS)}

    def column_width(index):
        width = default_width
        for column in columns:
            if int(column.get("min")) <= index <= int(column.get("max")):
                if column.get("hidden") in {"1", "true"}:
                    return 0
                width = float(column.get("width", str(default_width)))
        return math.floor(width * 7 + .5) * .75

    adjusted = []
    for number, row in rows.items():
        if row.get("hidden") in {"1", "true"}:
            continue
        headings = []
        required = float(row.get("ht", str(default_height)))
        for cell in row.findall("s:c", NS):
            text = value_of(cell, strings).strip()
            if not text or not printed(cell.get("r", "")):
                continue
            style = formats[int(cell.get("s", "0"))]
            alignment = style.find("s:alignment", NS)
            fill = fills[int(style.get("fillId", "0"))].find("s:patternFill", NS)
            border = borders[int(style.get("borderId", "0"))]
            if (alignment is None or alignment.get("wrapText") not in {"1", "true"}
                    or alignment.get("textRotation", "0") != "0"
                    or fill is None or fill.get("patternType") != "solid"
                    or not any(edge.get("style") for edge in border)):
                continue
            first, _ = _coordinate(cell.get("r"))
            last = first
            merged = next((m for m in merges if m[0] <= first <= m[2] and m[1] <= number <= m[3]), None)
            if merged:
                if merged[1] != merged[3]:
                    continue
                first, _, last, _ = merged
            font = fonts[int(style.get("fontId", "0"))]
            name = font.find("s:name", NS)
            if name is None or name.get("val", "").lower() not in {"arial", "calibri", "carlito", "liberation sans"}:
                continue
            font_size = font.find("s:sz", NS)
            size = float(font_size.get("val", "11")) if font_size is not None else 11
            metrics = pymupdf.Font("hebo" if font.find("s:b", NS) is not None else "helv")
            width = max(1, sum(column_width(c) for c in range(first, last + 1)) - 6)
            lines = 0
            for paragraph in text.split("\n"):
                line = ""
                lines += 1
                for word in paragraph.split():
                    candidate = (line + " " + word).strip()
                    if line and metrics.text_length(candidate, fontsize=size) > width:
                        lines += 1
                        line = word
                    else:
                        line = candidate
            headings.append(cell.get("r"))
            if lines > 1:
                required = max(required, lines * size * 1.25 + 3)
        height = float(row.get("ht", str(default_height)))
        extra = required - height
        if len(headings) < 2 or extra <= .1:
            continue
        # Do not cross a break, merged cell, styled line or occupied row.
        spacers = []
        for previous in range(number - 1, max(0, number - 5), -1):
            spacer = rows.get(previous)
            if previous in breaks or spacer is None or spacer.get("hidden") in {"1", "true"}:
                break
            if any(m[1] <= previous <= m[3] for m in merges):
                break
            if any(value_of(c, strings).strip() or any(
                    edge.get("style") for edge in borders[int(formats[int(c.get("s", "0"))].get("borderId", "0"))])
                    for c in spacer.findall("s:c", NS)):
                break
            available = max(0, float(spacer.get("ht", str(default_height))) - 3)
            spacers.append((spacer, available))
            if sum(amount for _, amount in spacers) >= extra:
                break
        if sum(amount for _, amount in spacers) + .001 < extra:
            continue
        row.set("ht", f"{required:g}")
        row.set("customHeight", "1")
        for spacer, available in spacers:
            used = min(available, extra)
            spacer.set("ht", f"{float(spacer.get('ht', str(default_height))) - used:g}")
            spacer.set("customHeight", "1")
            extra -= used
        adjusted.extend(headings)
    return adjusted
