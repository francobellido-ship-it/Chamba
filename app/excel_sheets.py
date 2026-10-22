"""Find numbered annex sheets belonging to the same saved certificate."""

import re
from dataclasses import replace

from app.pdf import CertificateError, CODE, MAX_PAGES


def certificate_layouts(data: bytes):
    from app.excel import _read_parts, _layout, xml, related, cell_value, NS, R
    from app.excel_ranges import printed_cell_checker
    parts = _read_parts(data)
    primary = _layout(parts)
    book = xml(parts["xl/workbook.xml"])
    strings = (["".join(item.itertext()) for item in xml(parts["xl/sharedStrings.xml"])]
               if "xl/sharedStrings.xml" in parts else [])

    def markers(layout):
        sheet = xml(parts[layout.sheet_path])
        printed = printed_cell_checker(book, sheet, layout.sheet_index, layout.sheet_name)
        text = " ".join(cell_value(cell, strings) for cell in sheet.findall("s:sheetData/s:row/s:c", NS)
                        if printed(cell.get("r", "")))
        codes = {re.sub(r"\s+", "", code.upper()) for code in CODE.findall(text)}
        pages = {(int(number), int(total)) for number, total in
                 re.findall(r"p[aá]gina\s*([0-9]+)\s*de\s*([0-9]+)", text, re.I)}
        return codes, pages

    codes, pages = markers(primary)
    totals = {total for number, total in pages}
    if len(codes) != 1 or len(totals) != 1 or not any(number == 1 for number, total in pages):
        return [primary]
    total = next(iter(totals))
    if not 1 <= total <= MAX_PAGES:
        raise CertificateError(f"El certificado debe tener entre 1 y {MAX_PAGES} páginas.")
    expected = set(range(1, total + 1))
    found = {number for number, _ in pages}
    if found == expected:
        return [replace(primary, page_numbers=tuple(sorted(found)))]
    numbered = [(min(found), replace(primary, page_numbers=tuple(sorted(found))))]
    targets = {rel.get("Id"): related("xl/workbook.xml", rel.get("Target", ""))
               for rel in xml(parts["xl/_rels/workbook.xml.rels"])
               if rel.get("TargetMode") != "External"}
    for index, sheet in enumerate(book.findall("s:sheets/s:sheet", NS)):
        if index == primary.sheet_index or sheet.get("state", "visible") != "visible":
            continue
        if not re.match(r"certificado(?:\s|[0-9])", sheet.get("name", "").strip(), re.I):
            continue
        if targets.get(sheet.get(f"{{{R}}}id")) not in parts:
            continue
        layout = _layout(parts, index)
        other_codes, other_pages = markers(layout)
        primary_series = re.fullmatch(r'(certificado(?:\s+qcp)?)\s*1',primary.sheet_name.strip(),re.I)
        other_series = re.fullmatch(r'(certificado(?:\s+qcp)?)\s*([0-9]+)',layout.sheet_name.strip(),re.I)
        continuation = (not other_codes and primary_series and other_series
                        and primary_series[1].casefold()==other_series[1].casefold()
                        and other_pages and int(other_series[2])==min(number for number,_ in other_pages))
        if (other_codes != codes and not continuation) or not other_pages or {count for _, count in other_pages} != {total}:
            continue
        numbers = {number for number, _ in other_pages}
        if numbers & found or not numbers <= expected:
            raise CertificateError("Las hojas del certificado tienen números de página duplicados o inconsistentes. Revisa su numeración.")
        found.update(numbers)
        numbered.append((min(numbers), replace(layout, page_numbers=tuple(sorted(numbers)))))
    if found != expected:
        missing = ", ".join(str(number) for number in sorted(expected - found))
        raise CertificateError(f"El certificado indica {total} páginas, pero faltan las hojas de las páginas {missing}. Inclúyelas con el mismo código y numeración para generar el PDF completo.")
    ordered = [layout for _, layout in sorted(numbered, key=lambda item: item[0])]
    if [number for layout in ordered for number in layout.page_numbers] != list(range(1, total + 1)):
        raise CertificateError("Las páginas de cada hoja deben ser consecutivas para unir el certificado en el orden correcto.")
    return ordered
