"""Make wrapped table headings readable without moving saved page boundaries."""

import math
import re

import pymupdf

from app.excel_ranges import NS


def fit_overflow_labels(sheet, styles, strings, value_of, printed, *, drawing_ranges=()):
    """Give overflowing labels their empty horizontal space without moving rows.

    Excel lets unmerged text flow across empty cells. Calc can cut it at the
    next occupied/merged cell, and table borders can pass through the label.
    Merge only compatible empty cells in the temporary copy and shrink within
    that span when necessary. Never overwrite data, merges or drawings.
    """
    from copy import deepcopy
    import xml.etree.ElementTree as ET
    if drawing_ranges is None:
        return []
    formats = styles.find('s:cellXfs', NS)
    fonts = styles.findall('s:fonts/s:font', NS)
    borders = styles.findall('s:borders/s:border', NS)
    properties = sheet.find('s:sheetFormatPr', NS)
    default = float(properties.get('defaultColWidth', '8.43')) if properties is not None else 8.43
    columns = sheet.findall('s:cols/s:col', NS)
    merges = sheet.find('s:mergeCells', NS)
    merged = [(*_coordinate(m.get('ref').split(':')[0]), *_coordinate(m.get('ref').split(':')[-1]))
              for m in merges] if merges is not None else []

    def width(column):
        value = default
        for item in columns:
            if int(item.get('min')) <= column <= int(item.get('max')):
                if item.get('hidden') in {'1','true'}:
                    return 0
                value = float(item.get('width', str(default)))
        return math.floor(value * 7 + 5) * .75

    def reference(column, row):
        letters = ''
        while column:
            column, digit = divmod(column-1, 26)
            letters = chr(65+digit) + letters
        return f'{letters}{row}'

    adjusted = []
    replacements = {}
    for row in sheet.findall('s:sheetData/s:row', NS):
        number = int(row.get('r'))
        cells = { _coordinate(c.get('r'))[0]: c for c in row.findall('s:c', NS) }
        for first, cell in cells.items():
            text = value_of(cell, strings).strip()
            if (cell.get('t') not in {'s','str','inlineStr'} or not text or '\n' in text
                    or not printed(cell.get('r')) or any(a<=first<=b and top<=number<=bottom for a,top,b,bottom in merged)):
                continue
            index = int(cell.get('s','0')); style = formats[index]
            alignment = style.find('s:alignment', NS)
            if alignment is not None and (alignment.get('wrapText') in {'1','true'}
                    or alignment.get('horizontal','left') not in {'left','general'}
                    or alignment.get('textRotation','0') != '0'):
                continue
            font = fonts[int(style.get('fontId','0'))]
            name = font.find('s:name', NS); size = font.find('s:sz', NS)
            if name is None or size is None or name.get('val','').casefold() not in {'arial','calibri','carlito','liberation sans'}:
                continue
            metrics = pymupdf.Font('hebo' if font.find('s:b',NS) is not None else 'helv')
            needed = metrics.text_length(text, fontsize=float(size.get('val')))+6
            if needed <= width(first):
                continue
            last = first
            bordered = any(edge.get('style') for edge in borders[int(style.get('borderId','0'))])
            for column in range(first+1, min(first+64, 16385)):
                other = cells.get(column)
                if (not printed(reference(column,number)) or width(column)==0
                        or any(a<=column<=b and top<=number<=bottom for a,top,b,bottom in merged)
                        or other is not None and (value_of(other,strings).strip() or other.find('s:f',NS) is not None)):
                    break
                other_style = formats[int(other.get('s','0'))] if other is not None else formats[0]
                other_bordered = any(edge.get('style') for edge in borders[int(other_style.get('borderId','0'))])
                if (bordered != other_bordered or bordered and
                        (style.get('fillId','0'),style.get('borderId','0')) != (other_style.get('fillId','0'),other_style.get('borderId','0'))):
                    break
                last = column
            if last == first or any(a<=last and b>=first and top<=number<=bottom for a,top,b,bottom in drawing_ranges):
                continue
            available = sum(width(c) for c in range(first,last+1))
            if not bordered and needed <= available:
                continue  # Plain Excel overflow already fits; preserve its cells.
            if merges is None:
                merges = ET.SubElement(sheet, f'{{{NS["s"]}}}mergeCells')
            ET.SubElement(merges, f'{{{NS["s"]}}}mergeCell', ref=f'{cell.get("r")}:{reference(last,number)}')
            merges.set('count',str(len(merges)));merged.append((first,number,last,number))
            if needed > available:
                if index not in replacements:
                    copy = deepcopy(style);target=copy.find('s:alignment',NS)
                    if target is None:target=ET.SubElement(copy,f'{{{NS["s"]}}}alignment')
                    target.set('shrinkToFit','1');copy.set('applyAlignment','1')
                    replacements[index]=len(formats);formats.append(copy)
                cell.set('s',str(replacements[index]))
            adjusted.append(cell.get('r'))
    formats.set('count',str(len(formats)))
    return adjusted


def wrap_table_lines(sheet, styles, strings, value_of, printed):
    """Calc needs wrapText to retain hard line breaks inside bordered cells."""
    from copy import deepcopy
    import xml.etree.ElementTree as ET
    formats=styles.find('s:cellXfs',NS)
    borders=styles.findall('s:borders/s:border',NS)
    replacements={}
    for cell in sheet.findall('s:sheetData/s:row/s:c',NS):
        if not printed(cell.get('r','')) or '\n' not in value_of(cell,strings):
            continue
        index=int(cell.get('s','0'));style=formats[index]
        alignment=style.find('s:alignment',NS)
        if (alignment is not None and (alignment.get('textRotation','0')!='0' or alignment.get('wrapText') in {'1','true'})
                or not any(edge.get('style')for edge in borders[int(style.get('borderId','0'))])):
            continue
        if index not in replacements:
            copy=deepcopy(style);target=copy.find('s:alignment',NS)
            if target is None:target=ET.SubElement(copy,f'{{{NS["s"]}}}alignment')
            target.set('wrapText','1');copy.set('applyAlignment','1')
            replacements[index]=len(formats);formats.append(copy)
        cell.set('s',str(replacements[index]))
    formats.set('count',str(len(formats)))


def fit_table_fonts(sheet, styles, strings, value_of, printed):
    """Keep bordered table text readable in fixed-scale worksheet exports."""
    from copy import deepcopy
    setup = sheet.find('s:pageSetup',NS)
    properties = sheet.find('s:sheetPr/s:pageSetUpPr',NS)
    if setup is None or properties is not None and properties.get('fitToPage') in {'1','true'}:
        return []
    scale = float(setup.get('scale','100')) / 100
    if not 0 < scale < 1:
        return []
    minimum = math.ceil(6.5 / scale * 2) / 2
    fonts = styles.find('s:fonts',NS)
    formats = styles.find('s:cellXfs',NS)
    borders = styles.findall('s:borders/s:border',NS)
    replacements, adjusted = {}, []
    for cell in sheet.findall('s:sheetData/s:row/s:c',NS):
        if not printed(cell.get('r','')) or not value_of(cell,strings).strip():
            continue
        index = int(cell.get('s','0'))
        style = formats[index]
        font = fonts[int(style.get('fontId','0'))]
        size = font.find('s:sz',NS)
        alignment = style.find('s:alignment',NS)
        if (size is None or float(size.get('val','11')) >= minimum
                or alignment is not None and alignment.get('textRotation','0')!='0'
                or not any(edge.get('style') for edge in borders[int(style.get('borderId','0'))])):
            continue
        if index not in replacements:
            new_font = deepcopy(font);new_font.find('s:sz',NS).set('val',f'{minimum:g}')
            new_style = deepcopy(style);new_style.set('fontId',str(len(fonts)))
            fonts.append(new_font)
            replacements[index] = len(formats);formats.append(new_style)
        cell.set('s',str(replacements[index]));adjusted.append(cell.get('r'))
    fonts.set('count',str(len(fonts)));formats.set('count',str(len(formats)))
    return adjusted


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


def fit_table_headings(sheet, styles, strings, value_of, printed, *, drawing_ranges=()):
    """Borrow nearby empty spacers; never grow the printed page.

    Calc does not auto-fit horizontally merged cells. Use conservative Arial
    metrics for the common Arial/Calibri headings, leaving fonts and data intact.
    Unsupported rotations, vertical merges and overlapping drawings retain their layout.
    """
    if drawing_ranges is None:
        return []
    fonts = styles.findall("s:fonts/s:font", NS)
    formats = styles.findall("s:cellXfs/s:xf", NS)
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
            border = borders[int(style.get("borderId", "0"))]
            if (alignment is None or (alignment.get("wrapText") not in {"1", "true"} and '\n' not in text)
                    or alignment.get("textRotation", "0") != "0"
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
            if any(m[1] <= previous <= m[3] and m[1]!=m[3] for m in merges):
                break
            cells = spacer.findall('s:c',NS)
            if any(any(edge.get('style') for edge in borders[int(formats[int(c.get('s','0'))].get('borderId','0'))]) for c in cells):
                break
            occupied = [c for c in cells if value_of(c,strings).strip()]
            if occupied:
                if len(occupied)==1:
                    caption_style = formats[int(occupied[0].get('s','0'))]
                    caption_font = fonts[int(caption_style.get('fontId','0'))]
                    caption_alignment = caption_style.find('s:alignment',NS)
                    if (caption_font.find('s:b',NS) is not None or
                            (caption_alignment is not None and caption_alignment.get('horizontal')=='center')):
                        continue  # A caption can move with the empty space above it.
                break
            available = max(0, float(spacer.get("ht", str(default_height))) - 3)
            spacers.append((spacer, available))
            if sum(amount for _, amount in spacers) >= extra:
                break
        if sum(amount for _, amount in spacers) + .001 < extra:
            for following in range(number+1,number+5):
                spacer = rows.get(following)
                if (following-1 in breaks or spacer is None or spacer.get('hidden') in {'1','true'}
                        or any(m[1]<=following<=m[3] and m[1]!=m[3] for m in merges)):
                    break
                if any(value_of(c,strings).strip() or any(edge.get('style') for edge in
                       borders[int(formats[int(c.get('s','0'))].get('borderId','0'))]) for c in spacer.findall('s:c',NS)):
                    break
                available = max(0,float(spacer.get('ht',str(default_height)))-3)
                spacers.append((spacer,available))
                if sum(amount for _,amount in spacers)>=extra:
                    break
        if sum(amount for _, amount in spacers) + .001 < extra:
            continue
        first_affected = min([number]+[int(spacer.get('r'))for spacer,_ in spacers])
        last_affected = max([number]+[int(spacer.get('r'))for spacer,_ in spacers])
        if any(top <= last_affected and bottom >= first_affected
               for _,top,_,bottom in drawing_ranges):
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
