"""Compose numbered worksheet pages without splitting their saved data blocks."""

import re
import math

from app.excel_ranges import NS, _bounds, column_number, print_names, printed_cell_checker
from app.pdf import CertificateError
import pymupdf

S = NS['s']
PAGE = re.compile(r'p[aá]gina\s*([0-9]+)\s*de\s*([0-9]+)', re.I)


def validate_page_labels(doc, numbers, total):
    for page, number in zip(doc, numbers, strict=True):
        labels = {(int(match[1]), int(match[2])) for match in PAGE.finditer(page.get_text())}
        if labels != {(number, total)}:
            raise CertificateError(f'La numeración de la página {number} está recortada o no coincide con la plantilla. No se emitió el certificado incompleto.')


def validate_table_text(doc):
    """Check table text against stroked borders, never background fills.

    Calc emits filled rectangles for charts and portions of merged cells.
    Treating those fills as grid lines invents cells through chart titles and
    stacked labels, rejecting readable templates that need no layout changes.
    """
    for index, page in enumerate(doc):
        cells = [pymupdf.Rect(cell) for table in page.find_tables(strategy="lines_strict").tables
                 for cell in table.cells if cell is not None]
        for block in page.get_text('dict', flags=pymupdf.TEXTFLAGS_DICT & ~pymupdf.TEXT_PRESERVE_IMAGES)['blocks']:
            for line in block.get('lines', []):
                for span in line['spans']:
                    if not span['text'].strip():
                        continue
                    box = pymupdf.Rect(span['bbox'])
                    center = (box.tl + box.br) / 2
                    containing = [cell for cell in cells if cell.contains(center)]
                    if not containing:
                        continue
                    if span['size'] < 6.5 or not any((cell + (-1, -1, 1, 1)).contains(box) for cell in containing):
                        raise CertificateError(f'Hay texto de tabla recortado o demasiado pequeño en la página {index + 1}. Revisa el tamaño de las celdas o distribuye ese cuadro en una página completa.')


def drawing_row_ranges(drawing, sheet):
    """Protect full drawing anchors, including one-cell images with an extent."""
    namespaces = {**NS, 'd': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
    heights = [float(row.get('ht', '15')) for row in sheet.findall('s:sheetData/s:row', NS)]
    minimum = min([15] + [height for height in heights if height > 0])
    ranges = []
    for anchor in drawing:
        start, end = anchor.find('d:from', namespaces), anchor.find('d:to', namespaces)
        if start is None:
            return None  # Unknown/absolute placement cannot be safely cropped.
        left = int(start.find('d:col', namespaces).text) + 1
        top = int(start.find('d:row', namespaces).text) + 1
        if end is not None:
            right = int(end.find('d:col', namespaces).text) + 1
            bottom = int(end.find('d:row', namespaces).text) + 1
        else:
            extent = anchor.find('d:ext', namespaces)
            if extent is None:
                return None
            # A conservative extent is sufficient for protecting row space;
            # original anchors, image bytes and transforms remain unchanged.
            right = 16384
            bottom = top + math.ceil(float(extent.get('cy', '0')) / 12700 / minimum) + 1
        ranges.append((left, top, right, bottom))
    return ranges


def _cell(reference):
    match = re.fullmatch(r'([A-Z]+)([0-9]+)', reference or '')
    return (column_number(match[1]), int(match[2])) if match else None


def _letters(number):
    result = ''
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def compose_numbered_pages(workbook, sheet, sheet_name, styles, strings, value_of, drawing_ranges=()):
    """Use complete page labels immediately below repeated title rows.

    Footer labels, unknown ranges, existing disjoint print areas and unrelated
    templates keep their native layout. No cell, saved value or drawing moves
    between logical pages. Only demonstrably empty spacers can lose height.
    """
    if drawing_ranges is None:
        return None
    names = print_names(workbook, 0)
    area, titles = names.get('Print_Area'), names.get('Print_Titles')
    if area is None or titles is None:
        return None
    ranges = _bounds(area.text or '', sheet_name)
    repeats = _bounds(titles.text or '', sheet_name)
    if not ranges or len(ranges) != 1 or not repeats or len(repeats) != 1:
        return None
    left, first, right, last = ranges[0]
    title_left, title_first, title_right, title_last = repeats[0]
    if title_left != 1 or title_right != 16384 or title_first != first:
        return None  # Only whole repeated rows; repeated columns need another layout.
    rows = {int(r.get('r')): r for r in sheet.findall('s:sheetData/s:row', NS)}
    printed = printed_cell_checker(workbook, sheet, 0, sheet_name)
    labels, meaningful, protected = [], set(), set(range(first, title_last + 1))
    formats = styles.findall('s:cellXfs/s:xf', NS)
    borders = styles.findall('s:borders/s:border', NS)
    fills = styles.findall('s:fills/s:fill', NS)
    values = {}
    for number, row in rows.items():
        if row.get('hidden') in {'1', 'true'}:
            continue
        for cell in row.findall('s:c', NS):
            position = _cell(cell.get('r'))
            if (not position or not left <= position[0] <= right or not first <= number <= last
                    or not printed(cell.get('r'))):
                continue
            text = value_of(cell, strings).strip()
            values[cell.get('r')] = text
            if text:
                meaningful.add(number)
                match = PAGE.fullmatch(text)
                if match:
                    labels.append((int(match[1]), int(match[2]), number, position[0]))
            style = formats[int(cell.get('s', '0'))]
            border = borders[int(style.get('borderId', '0'))]
            fill = fills[int(style.get('fillId', '0'))].find('s:patternFill', NS)
            if any(edge.get('style') for edge in border) or (fill is not None and fill.get('patternType') not in {None, 'none'}):
                meaningful.add(number)
    labels.sort()
    if len(labels) < 2 or {total for _, total, _, _ in labels} != {len(labels)}:
        return None
    if [n for n, _, _, _ in labels] != list(range(1, len(labels) + 1)):
        return None
    starts = [row for _, _, row, _ in labels]
    if starts != sorted(set(starts)) or len({col for _, _, _, col in labels}) != 1:
        return None
    if not title_last < starts[0] <= title_last + 3:
        return None  # A page number near the bottom is a footer, not a page start.
    if sheet.find('s:colBreaks/s:brk', NS) is not None:
        return None  # An explicit horizontal layout is not a vertical page series.
    # Never replace explicit saved breaks that disagree with the page labels.
    existing = {int(brk.get('id')) for brk in sheet.findall('s:rowBreaks/s:brk', NS)}
    if existing and existing != {row - 1 for row in starts[1:]}:
        return None
    merges = []
    for merge in sheet.findall('s:mergeCells/s:mergeCell', NS):
        refs = merge.get('ref').split(':')
        a, b = _cell(refs[0]), _cell(refs[-1])
        if not a or not b or b[0] < left or a[0] > right:
            continue
        merges.append((a[1], b[1]))
        if a[1] != b[1]:
            protected.update(range(a[1], b[1] + 1))
        if values.get(refs[0]) or any(row in meaningful for row in range(a[1], b[1] + 1)):
            meaningful.update(range(a[1], b[1] + 1))
    for draw_left, top, draw_right, bottom in drawing_ranges:
        if draw_right >= left and draw_left <= right:
            protected.update(range(max(first, top), min(last, bottom) + 1))
            meaningful.update(range(max(first, top), min(last, bottom) + 1))
    boundaries = starts[1:]
    for top, bottom in merges:
        if any(top < boundary <= bottom for boundary in boundaries):
            raise CertificateError('Una celda combinada cruza la numeración de páginas. Revisa ese bloque para conservar toda su información.')
    for draw_left, top, draw_right, bottom in drawing_ranges:
        if draw_right >= left and draw_left <= right and any(top < boundary <= bottom for boundary in boundaries):
            raise CertificateError('Una imagen o cuadro de texto cruza las páginas previstas del certificado. Revisa su anclaje antes de imprimir.')
    segments = []
    for top, bottom in zip([first] + starts[1:], [row - 1 for row in starts[1:]] + [last]):
        occupied = [row for row in meaningful if top <= row <= bottom]
        segments.append((top, max(occupied)))
    # Reduce excessive empty space inside a logical page, never inside a merged
    # paragraph, a table, a drawing, the repeated title or a hidden row.
    properties = sheet.find('s:sheetFormatPr', NS)
    default_height = float(properties.get('defaultRowHeight', '15')) if properties is not None else 15
    for top, bottom in segments:
        run = []
        def compact():
            if not run:
                return
            heights = [float(rows[n].get('ht', str(default_height))) for n in run]
            total = sum(heights)
            if total <= 36:
                return
            for number, height in zip(run, heights):
                rows[number].set('ht', f'{height * 24 / total:.6f}')
                rows[number].set('customHeight', '1')
        for number in range(top, bottom + 2):
            row = rows.get(number)
            if (number <= bottom and row is not None and number not in meaningful | protected
                    and row.get('hidden') not in {'1', 'true'}):
                run.append(number)
            else:
                compact()
                run = []
    name = sheet_name.replace("'", "''")
    area.text = ','.join(f"'{name}'!${_letters(left)}${top}:${_letters(right)}${bottom}" for top, bottom in segments)
    from xml.etree import ElementTree as ET
    setup = sheet.find('s:pageSetup', NS)
    if setup is None:
        setup = ET.SubElement(sheet, f'{{{S}}}pageSetup')
    setup.set('fitToWidth', '1')
    setup.set('fitToHeight', '1')
    properties = sheet.find('s:sheetPr', NS)
    if properties is None:
        properties = ET.Element(f'{{{S}}}sheetPr')
        sheet.insert(0, properties)
    page = properties.find('s:pageSetUpPr', NS)
    if page is None:
        page = ET.SubElement(properties, f'{{{S}}}pageSetUpPr')
    page.set('fitToPage', '1')
    for tag in ['rowBreaks', 'colBreaks']:
        breaks = sheet.find('s:' + tag, NS)
        if breaks is not None:
            sheet.remove(breaks)
    options = sheet.find('s:printOptions', NS)
    if options is None:
        options = ET.Element(f'{{{S}}}printOptions')
        margins = sheet.find('s:pageMargins', NS)
        sheet.insert(list(sheet).index(margins) if margins is not None else len(sheet), options)
    options.set('verticalCentered', '0')
    return tuple(segments)
