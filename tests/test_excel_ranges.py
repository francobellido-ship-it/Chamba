"""Check only saved, printed results; never hide an error that will be visible."""

from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

import pytest

from app.excel import NS, S, _prepare_workbook
from app.excel_ranges import _bounds
from app.pdf import CertificateError
from test_excel import workbook_bytes, rewrite


@pytest.mark.parametrize("outside,uncached", [(True, False), (True, True), (False, False), (False, True)])
def test_error_or_missing_result_outside_print_area_does_not_block_printing(outside, uncached):
    def change(parts):
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        if outside:
            row = root.find("s:sheetData/s:row[@r='63']", NS)
            cell = ET.SubElement(row, f"{{{S}}}c", r="CL63", t="e")
            ET.SubElement(cell, f"{{{S}}}f").text = "#REF!-#REF!"
            if not uncached:
                ET.SubElement(cell, f"{{{S}}}v").text = "#REF!"
        else:
            cell = root.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']", NS)
            cell.set("t", "e")
            if uncached:
                cell.remove(cell.find("s:v", NS))
            else:
                cell.find("s:v", NS).text = "#REF!"
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    data = rewrite(workbook_bytes(), change)
    if not outside:
        with pytest.raises(CertificateError, match="D5"):
            _prepare_workbook(data, False)
        return
    prepared, *_ = _prepare_workbook(data, False)
    with zipfile.ZipFile(BytesIO(prepared)) as source:
        sheet = ET.fromstring(source.read("xl/worksheets/sheet2.xml"))
        cell = sheet.find("s:sheetData/s:row[@r='63']/s:c[@r='CL63']", NS)
        assert cell.find("s:f", NS) is None
        assert sheet.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']/s:v", NS).text == "17.5"


def test_hidden_columns_do_not_block_but_keep_their_visibility():
    def change(parts):
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        col = next(c for c in root.find("s:cols", NS) if c.get("min") == "4")
        col.set("hidden", "1")
        root.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']", NS).remove(
            root.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']/s:v", NS))
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    prepared, *_ = _prepare_workbook(rewrite(workbook_bytes(), change), False)
    with zipfile.ZipFile(BytesIO(prepared)) as source:
        root = ET.fromstring(source.read("xl/worksheets/sheet2.xml"))
        assert root.find("s:cols/s:col[@min='4']", NS).get("hidden") == "1"


@pytest.mark.parametrize("area,titles,should_fail", [
    ("'Certificado'!$A$10:$L$70", "'Certificado'!$1:$5", True),
    ("'Certificado'!$A$10:$L$70,'Certificado'!$D$5", None, True),
    ("OFFSET(Certificado!A1,0,0,70,12)", None, True),
    ("'Certificado'!$A$10:$L$70", None, False),
])
def test_print_titles_union_and_unknown_ranges_never_hide_printed_errors(area, titles, should_fail):
    def change(parts):
        book = ET.fromstring(parts["xl/workbook.xml"])
        names = book.find("s:definedNames", NS)
        for name in list(names):
            if name.get("name") == "_xlnm.Print_Area":
                name.text = area
            elif name.get("name") == "_xlnm.Print_Titles":
                if titles:
                    name.text = titles
                else:
                    names.remove(name)
        parts["xl/workbook.xml"] = ET.tostring(book)
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        cell = root.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']", NS)
        cell.remove(cell.find("s:f", NS))  # Saved errors without formulas also appear in print.
        cell.set("t", "e")
        cell.find("s:v", NS).text = "#REF!"
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    data = rewrite(workbook_bytes(), change)
    if should_fail:
        with pytest.raises(CertificateError, match=r"D5.*#REF!"):
            _prepare_workbook(data, False)
    else:
        _prepare_workbook(data, False)


def test_standard_print_area_wins_over_duplicate_legacy_alias():
    def change(parts):
        book = ET.fromstring(parts["xl/workbook.xml"])
        names = book.find("s:definedNames", NS)
        ET.SubElement(names, f"{{{S}}}definedName", name="Print_Area", localSheetId="1").text = "'Certificado'!$A$1:$C$70"
        parts["xl/workbook.xml"] = ET.tostring(book)
    prepared, *_ = _prepare_workbook(rewrite(workbook_bytes(), change), False)
    with zipfile.ZipFile(BytesIO(prepared)) as source:
        book = ET.fromstring(source.read("xl/workbook.xml"))
        names = [n for n in book.findall("s:definedNames/s:definedName", NS) if n.get("name") == "_xlnm.Print_Area"]
        assert len(names) == 1 and names[0].text == "'Certificado'!$A$1:$L$70"


def test_range_parser_handles_sheet_names_with_commas_and_quotes():
    assert _bounds("'Certif, O''Brien'!$A$1:$AN$97", "Certif, O'Brien") == [(1, 1, 40, 97)]
    assert _bounds("'Other'!$A$1:$AN$97", "Certificado") is None
