"""Identify saved print ranges without evaluating workbook formulas."""

import re

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column_number(letters: str) -> int:
    result = 0
    for character in letters:
        result = result * 26 + ord(character.upper()) - 64
    return result


def print_names(workbook, sheet_index: int):
    """Prefer standard names over duplicate aliases saved by other editors."""
    selected = {}
    for name in workbook.findall("s:definedNames/s:definedName", NS):
        kind = name.get("name", "").removeprefix("_xlnm.")
        if name.get("localSheetId") == str(sheet_index) and kind in {"Print_Area", "Print_Titles"}:
            if kind not in selected or name.get("name", "").startswith("_xlnm."):
                selected[kind] = name
    return selected


_REF = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
_END = rf"(?:{_REF}|\$?[A-Za-z]{{1,3}}|\$?[1-9][0-9]{{0,6}})"
_RANGE = re.compile(
    rf"\s*(?:(?:'(?P<quoted>(?:[^']|'')+)'|(?P<plain>[^'!,]+))!)?"
    rf"(?P<first>{_END})(?::(?P<last>{_END}))?\s*(?:,|$)"
)


def _bounds(text: str, sheet_name: str):
    ranges, position = [], 0
    while position < len(text):
        match = _RANGE.match(text, position)
        if match is None:
            return None
        sheet = match["quoted"].replace("''", "'") if match["quoted"] else match["plain"]
        if sheet is not None and sheet.casefold() != sheet_name.casefold():
            return None
        ends = [value.replace("$", "").upper() for value in (match["first"], match["last"] or match["first"])]
        coordinates = []
        for end in ends:
            cell = re.fullmatch(r"([A-Z]+)([0-9]+)", end)
            if cell:
                coordinates.append((column_number(cell[1]), int(cell[2])))
            elif end.isdigit():
                coordinates.append((None, int(end)))
            else:
                coordinates.append((column_number(end), None))
        (left, top), (right, bottom) = coordinates
        if (left is None) != (right is None) or (top is None) != (bottom is None):
            return None
        left, top, right, bottom = left or 1, top or 1, right or 16384, bottom or 1048576
        if not (1 <= left <= right <= 16384 and 1 <= top <= bottom <= 1048576):
            return None
        ranges.append((left, top, right, bottom))
        position = match.end()
    return ranges or None


def printed_cell_checker(workbook, sheet, sheet_index: int, sheet_name: str):
    names = print_names(workbook, sheet_index)
    area = names.get("Print_Area")
    ranges = _bounds(area.text or "", sheet_name) if area is not None else None
    titles = names.get("Print_Titles")
    if ranges is not None and titles is not None:
        title_ranges = _bounds(titles.text or "", sheet_name)
        ranges = ranges + title_ranges if title_ranges is not None else None
    hidden_rows = {int(row.get("r")) for row in sheet.findall("s:sheetData/s:row", NS)
                   if row.get("hidden") in {"1", "true"}}
    hidden_columns = [(int(col.get("min")), int(col.get("max")))
                      for col in sheet.findall("s:cols/s:col", NS) if col.get("hidden") in {"1", "true"}]

    def printed(reference: str) -> bool:
        match = re.fullmatch(r"([A-Z]+)([0-9]+)", reference)
        if match is None:
            return True  # Unknown references must not bypass validation.
        col, row = column_number(match[1]), int(match[2])
        if row in hidden_rows or any(left <= col <= right for left, right in hidden_columns):
            return False
        # Unknown/dynamic print ranges fall back to checking all visible cells.
        return ranges is None or any(left <= col <= right and top <= row <= bottom
                                     for left, top, right, bottom in ranges)

    return printed
