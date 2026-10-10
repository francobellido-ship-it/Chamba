"""Synthetic workbooks only: no customer certificates or signature assets."""

from io import BytesIO
import subprocess
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.worksheet.pagebreak import Break
from PIL import Image
import pymupdf
import pytest
import zxingcpp

from app.excel import S, convert_excel, _prepare_workbook, inspect_excel
from app.main import create_app
from app.pdf import CertificateError, prepare_pdf


def image_bytes(size, color):
    buffer = BytesIO()
    Image.new("RGB", size, color).save(buffer, "PNG")
    return buffer.getvalue()


def rewrite(data, changes):
    result = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as source, zipfile.ZipFile(result, "w", zipfile.ZIP_DEFLATED) as output:
        parts = {name: source.read(name) for name in source.namelist()}
        changes(parts)
        for name, content in parts.items():
            output.writestr(name, content)
    return result.getvalue()


def workbook_bytes(code="QCP-2026-0001"):
    workbook = Workbook()
    other = workbook.active
    other.title = "Registro"
    other["A1"] = "OTHER SHEET MUST NOT BE PRINTED"
    other["A2"] = "=1/0"  # Irrelevant uncached formula outside the print sheet.
    sheet = workbook.create_sheet("Certificado")
    sheet["A1"] = "FICTICIO - SIN VALIDEZ"
    sheet["H3"] = '=Registro!A1'
    sheet["A5"] = "Resultado guardado:"
    sheet["D5"] = "=1+1"
    sheet["D5"].number_format = "0.0"
    sheet["D20"] = "Escanee este QR"
    sheet["D21"] = "Para consultar el certificado"
    sheet["J20"] = "Autorizado y firmador por:"
    sheet["J21"] = "PERSONA FICTICIA"
    sheet["A40"] = "CONTENIDO DE LA SEGUNDA PAGINA"
    for column in "ABCDEFGHIJKL":
        sheet.column_dimensions[column].width = 8.3
    for index in range(1, 71):
        sheet.row_dimensions[index].height = 14
    sheet.print_area = "A1:L70"
    sheet.print_title_rows = "1:3"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.row_breaks.append(Break(id=36))
    for anchor, size, color in [("F20", (60, 60), "red"), ("H20", (100, 45), "blue"), ("F45", (60, 60), "green")]:
        sheet.add_image(ExcelImage(BytesIO(image_bytes(size, color))), anchor)
    output = BytesIO()
    workbook.save(output)

    def cached_values(parts):
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        for cell in root.findall(f".//{{{S}}}c"):
            value = cell.find(f"{{{S}}}v")
            if cell.get("r") == "H3":
                cell.set("t", "str")
                value.text = code
            if cell.get("r") == "D5":
                value.text = "17.5"  # Saved result must not be recalculated as 2.
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    return rewrite(output.getvalue(), cached_values)


def qr_images(doc):
    decoded = []
    for info in doc[0].get_image_info(xrefs=True):
        if pymupdf.Rect(info["bbox"]).get_area() > doc[0].rect.get_area() * 0.9:
            continue
        # Render transparency against the page, as a scanner sees the PDF.
        pixmap = doc[0].get_pixmap(matrix=pymupdf.Matrix(3, 3), clip=pymupdf.Rect(info["bbox"]))
        image = Image.open(BytesIO(pixmap.tobytes("png")))
        barcode = zxingcpp.read_barcode(image)
        if barcode:
            decoded.append((barcode.text, pymupdf.Rect(info["bbox"])))
    return decoded


def test_excel_batch_has_distinct_qrs_and_preserves_layout_and_saved_results(tmp_path):
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        urls = []
        for index in range(2):
            response = client.post("/api/certificates", files={
                "excel": (f"certificate-{index}.xlsx", workbook_bytes(f"QCP-2026-000{index + 1}"))},
                data={"tipoCertificado": "acreditado"})
            assert response.status_code == 201, response.text
            result = response.json()
            assert result["filename"] == f"certificate-{index}.pdf"
            assert result["sourceType"] == "xlsx"
            assert result["signatureIncluded"] is True
            assert result["stampPages"] == [1]
            assert any("sellos" in warning for warning in result["warnings"])
            with pymupdf.open(stream=client.get(result["downloadUrl"]).content, filetype="pdf") as doc:
                assert len(doc) == 2
                text = "\n".join(page.get_text() for page in doc)
                assert "17,5" in text
                assert "OTHER SHEET" not in text
                assert "SEGUNDA PAGINA" in doc[1].get_text()
                assert result["finalCode"] in text
                assert len(doc[0].get_image_info()) == 3  # Official background, signature and QR.
                assert len(doc[1].get_image_info()) == 1  # Background only; no seal or QR.
                decoded = qr_images(doc)
                assert [url for url, area in decoded] == [result["verificationUrl"]]
                area = decoded[0][1]
                guide = doc[0].search_for("Escanee este QR")[0]
                assert area.x1 < guide.x0
                assert abs(area.y0 - guide.y0) < 3
                urls.append(result["verificationUrl"])
            assert client.get(f"/certificados/{result['id']}").status_code == 200
        assert urls[0] != urls[1]


def test_optional_signature_replaces_existing_image_and_corrects_code(tmp_path):
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        response = client.post("/api/certificates", files={
            "excel": ("certificate.xlsx", workbook_bytes()),
            "firma": ("signature.png", image_bytes((120, 40), "purple"))},
            data={"tipoCertificado": "acreditado", "codigoCorrecto": "QCP-2026-0099"})
        assert response.status_code == 201, response.text
        result = response.json()
        assert result["signatureIncluded"]
        assert result["finalCode"] == "QCP-2026-0099"
        assert result["replacements"] == 2  # Repeated print-title code on both pages.
        with pymupdf.open(stream=client.get(result["downloadUrl"]).content, filetype="pdf") as doc:
            sizes = [(info["width"], info["height"]) for info in doc[0].get_image_info()]
            assert (120, 40) in sizes
            assert (100, 45) not in sizes
            assert len(doc) == 2


@pytest.mark.parametrize("cell", ["D5", "H3"])
def test_missing_cached_formula_results_are_rejected(cell):
    def remove_cache(parts):
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        target = root.find(f".//{{{S}}}c[@r='{cell}']")
        target.remove(target.find(f"{{{S}}}v"))
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    with pytest.raises(CertificateError, match="resultados guardados"):
        convert_excel(rewrite(workbook_bytes(), remove_cache))


@pytest.mark.parametrize("old,new,error", [
    (b"Certificado", b"OtraHoja", "hoja llamada Certificado"),
])
def test_missing_template_markers_are_actionable(old, new, error):
    def replace(parts):
        for name in ["xl/workbook.xml", "xl/worksheets/sheet2.xml"]:
            parts[name] = parts[name].replace(old, new)
    with pytest.raises(CertificateError, match=error):
        convert_excel(rewrite(workbook_bytes(), replace))


def test_data_connections_are_removed_from_temporary_workbook():
    def add_connection(parts):
        parts["xl/connections.xml"] = b'<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'
    prepared, _, _, _, _ = _prepare_workbook(rewrite(workbook_bytes(), add_connection), False)
    with zipfile.ZipFile(BytesIO(prepared)) as archive:
        assert "xl/connections.xml" not in archive.namelist()


def test_macro_workbook_is_rejected():
    def add_macro(parts):
        parts["xl/vbaProject.bin"] = b"not executed"
    with pytest.raises(CertificateError, match="sin macros"):
        convert_excel(rewrite(workbook_bytes(), add_macro))


def test_converter_timeout_is_actionable(monkeypatch):
    def timeout(*args, **kwargs):
        raise subprocess.TimeoutExpired("soffice", 120)
    monkeypatch.setattr("app.excel.subprocess.run", timeout)
    with pytest.raises(CertificateError, match="tardó demasiado"):
        convert_excel(workbook_bytes())


def test_missing_or_ambiguous_upload_is_rejected(tmp_path):
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        assert client.post("/api/certificates").status_code == 422
        response = client.post("/api/certificates", files={"excel": ("invalid.xlsx", b"invalid")},
                               data={"tipoCertificado": "acreditado"})
        assert response.status_code == 422
        assert "Excel .xlsx válido" in response.json()["detail"]
        assert client.post("/api/certificates", files={
            "excel": ("one.xlsx", b"test"), "pdf": ("two.pdf", b"test")}).status_code == 422
        assert not list(tmp_path.iterdir())


def test_template_collision_fails_without_adding_a_page():
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_text((200, 300), "Escanee este QR")
        page.insert_text((130, 330), "Existing content")
        with pytest.raises(CertificateError, match="espacio libre"):
            prepare_pdf(doc.tobytes(), "", "https://example.com/test", template=True)


def test_template_without_visible_guide_fails():
    with pymupdf.open() as doc:
        doc.new_page().insert_text((50, 100), "No printed guide")
        with pytest.raises(CertificateError, match="identificar el espacio para el QR"):
            prepare_pdf(doc.tobytes(), "", "https://example.com/test", template=True)


def test_template_qr_only_on_main_page_when_guide_repeats():
    with pymupdf.open() as source:
        for _ in range(2):
            source.new_page().insert_text((200, 300), "Escanee este QR")
        result = prepare_pdf(source.tobytes(), "", "https://example.com/test", template=True)
    assert result.stamp_pages == [1]
    with pymupdf.open(stream=result.data, filetype="pdf") as doc:
        assert len(doc[0].get_images()) == 1
        assert not doc[1].get_images()


def test_wrapped_guide_is_one_slot_and_keeps_qr_beside_it():
    with pymupdf.open() as source:
        page = source.new_page()
        page.insert_textbox(pymupdf.Rect(200, 300, 255, 370), "Escanee este QR para consultar", fontsize=10)
        result = prepare_pdf(source.tobytes(), "", "https://example.com/test", template=True)
    with pymupdf.open(stream=result.data, filetype="pdf") as doc:
        decoded = qr_images(doc)
        assert [url for url, area in decoded] == ["https://example.com/test"]
        assert decoded[0][1].x1 < 200
        assert len(doc) == 1


def test_multiple_actual_qr_guides_are_still_rejected():
    with pymupdf.open() as source:
        page = source.new_page()
        page.insert_text((200, 300), "Escanee este QR")
        page.insert_text((200, 500), "Escanee este QR")
        with pytest.raises(CertificateError, match="varios espacios"):
            prepare_pdf(source.tobytes(), "", "https://example.com/test", template=True)


def two_certificate_sheets(data):
    def adapt(parts):
        workbook = ET.fromstring(parts["xl/workbook.xml"])
        sheets = workbook.find(f"{{{S}}}sheets")
        sheets[0].set("name", "CERTIFICADO")
        sheets[1].set("name", "CERTIFICADO QCP")
        workbook.remove(workbook.find(f"{{{S}}}definedNames"))
        parts["xl/workbook.xml"] = ET.tostring(workbook)
        sheet = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        setup = sheet.find(f"{{{S}}}pageSetup")
        setup.attrib.pop("paperSize", None)
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
    return rewrite(data, adapt)


def test_correct_sheet_selected_without_print_area_and_no_google_write(tmp_path):
    data = two_certificate_sheets(workbook_bytes())
    layout = inspect_excel(data)
    assert layout.sheet_name == "CERTIFICADO QCP"
    assert not layout.preserve_format
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        check = client.post("/api/excel/inspect", files={"excel": ("test.xlsx", data)})
        assert check.status_code == 200
        assert check.json() == {"sheetName": "CERTIFICADO QCP", "preserveFormat": False, "needsBackground": True}
        assert not list(tmp_path.iterdir())
        response = client.post("/api/certificates", files={"excel": ("test.xlsx", data)}, data={"tipoCertificado": "no_acreditado"})
        assert response.status_code == 201, response.text
        assert response.json()["sheetName"] == "CERTIFICADO QCP"
        with pymupdf.open(stream=client.get(response.json()["downloadUrl"]).content, filetype="pdf") as doc:
            assert len(doc) == 2
            assert all(abs(page.rect.width - 595.276) < 1 for page in doc)
            assert "OTHER SHEET" not in "\n".join(page.get_text() for page in doc)
            assert [url for url, _ in qr_images(doc)] == [response.json()["verificationUrl"]]
