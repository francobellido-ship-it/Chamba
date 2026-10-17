"""Auxiliary errors continue with warnings; printed measurement errors block."""

from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest

from app.excel import NS, S, _prepare_workbook, xml
from app.pdf import CertificateError
from test_excel import workbook_bytes, rewrite


def header_reference_workbook(*, variant="auxiliary"):
    def change(parts):
        styles = xml(parts["xl/styles.xml"])
        fonts = styles.find("s:fonts", NS)
        font_id = len(fonts)
        font = ET.SubElement(fonts, f"{{{S}}}font")
        ET.SubElement(font, f"{{{S}}}sz", val="10")
        ET.SubElement(font, f"{{{S}}}color", rgb="FF000000" if variant == "black" else "FFFFFFFF")
        fonts.set("count", str(len(fonts)))
        formats = styles.find("s:cellXfs", NS)
        style_id = len(formats)
        style = ET.SubElement(formats, f"{{{S}}}xf", fontId=str(font_id), fillId="0", borderId="0", numFmtId="0", xfId="0", applyFont="1", applyProtection="1")
        ET.SubElement(style, f"{{{S}}}protection", hidden="0" if variant == "unprotected" else "1")
        if variant == "colored_fill":
            fills = styles.find("s:fills", NS)
            style.set("fillId", str(len(fills)))
            fill = ET.SubElement(fills, f"{{{S}}}fill")
            pattern = ET.SubElement(fill, f"{{{S}}}patternFill", patternType="solid")
            ET.SubElement(pattern, f"{{{S}}}fgColor", rgb="FF006600")
            fills.set("count", str(len(fills)))
        if variant == "table_border":
            borders = styles.find("s:borders", NS)
            style.set("borderId", str(len(borders)))
            border = ET.SubElement(borders, f"{{{S}}}border")
            ET.SubElement(border, f"{{{S}}}bottom", style="thin")
            borders.set("count", str(len(borders)))
        formats.set("count", str(len(formats)))
        parts["xl/styles.xml"] = ET.tostring(styles)
        sheet = xml(parts["xl/worksheets/sheet2.xml"])
        for reference, text in [("A11", "CERTIFICADO DE CALIBRACIÓN QCP-2026-0001"), ("A13", "Página 2 de 2")]:
            if variant == "no_title" and reference == "A11":
                continue
            if variant == "no_page_number" and reference == "A13":
                continue
            row = sheet.find(f"s:sheetData/s:row[@r='{reference[1:]}']", NS)
            cell = ET.SubElement(row, f"{{{S}}}c", r=reference, t="inlineStr")
            ET.SubElement(ET.SubElement(cell, f"{{{S}}}is"), f"{{{S}}}t").text = text
        row = sheet.find("s:sheetData/s:row[@r='12']", NS)
        cell = ET.SubElement(row, f"{{{S}}}c", r="H12", t="e", s=str(style_id))
        ET.SubElement(cell, f"{{{S}}}f").text = "#REF!"
        ET.SubElement(cell, f"{{{S}}}v").text = "#DIV/0!" if variant == "other_error" else "#REF!"
        if variant == "missing_result":
            cell.remove(cell.find("s:v", NS))
            cell.attrib.pop("t")
        if variant == "numeric_format":
            style.set("numFmtId", "2")
            parts["xl/styles.xml"] = ET.tostring(styles)
        if variant in {"conditional", "unknown_conditional"}:
            area = "H12:J12" if variant == "conditional" else "UnknownName"
            ET.SubElement(sheet, f"{{{S}}}conditionalFormatting", sqref=area)
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
    return rewrite(workbook_bytes(), change)


@pytest.mark.parametrize("variant", ["auxiliary", "no_title", "no_page_number", "other_error", "missing_result"])
def test_auxiliary_errors_do_not_block_and_always_produce_a_warning(variant):
    original = header_reference_workbook(variant=variant)
    prepared, *_, warnings = _prepare_workbook(original, False)
    with ZipFile(BytesIO(prepared)) as archive, ZipFile(BytesIO(original)) as source:
        sheet = xml(archive.read("xl/worksheets/sheet2.xml"))
        cell = sheet.find("s:sheetData/s:row[@r='12']/s:c[@r='H12']", NS)
        assert cell.find("s:v", NS) is None and cell.find("s:f", NS) is None
        assert cell.get("s") is not None
        assert sheet.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']/s:v", NS).text == "17.5"
        assert b"#REF!" in source.read("xl/worksheets/sheet2.xml")
    assert any("H12" in warning and "errores auxiliares" in warning for warning in warnings)


@pytest.mark.parametrize("variant", ["black", "colored_fill", "table_border", "unprotected", "numeric_format", "conditional", "unknown_conditional"])
def test_visible_or_ambiguous_errors_are_still_rejected(variant):
    with pytest.raises(CertificateError, match="H12"):
        _prepare_workbook(header_reference_workbook(variant=variant), False)


def test_missing_certificate_code_is_a_critical_error():
    from app.documents import prepare_document
    with pytest.raises(CertificateError, match="código del certificado"):
        prepare_document(workbook_bytes(code="SIN CODIGO"), "xlsx", "", "https://example.com/test", certificate_type="no_acreditado")
