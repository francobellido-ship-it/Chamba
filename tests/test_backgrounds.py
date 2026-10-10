"""Synthetic letterheads test page coverage, selection and publication failures."""

from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace
from io import BytesIO
import hashlib
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
from PIL import Image, ImageDraw
import pymupdf
import pytest
import zxingcpp

import app.backgrounds as backgrounds
from app.documents import prepare_document
from app.excel import ConvertedExcel, S, _prepare_workbook
from app.main import create_app
from app.pdf import CertificateError
from test_drive import FakeDrive, ROOT
from test_excel import workbook_bytes, rewrite


@pytest.fixture
def synthetic_backgrounds(tmp_path, monkeypatch):
    specifications = {}
    for kind, color in (("acreditado", "#dce8ef"), ("no_acreditado", "#e4efdc")):
        image = Image.new("RGB", (248, 351), color)
        ImageDraw.Draw(image).text((10, 15), "PRUEBA SIN VALIDEZ " + kind, fill="black")
        output = BytesIO()
        image.save(output, "PNG")
        data = output.getvalue()
        path = tmp_path / (kind + ".png")
        path.write_bytes(data)
        specifications[kind] = (kind, path, hashlib.sha256(data).hexdigest())
    monkeypatch.setattr(backgrounds, "TEMPLATES", specifications)
    backgrounds.load_letterhead.cache_clear()
    yield specifications
    backgrounds.load_letterhead.cache_clear()


def page_with_background(doc, selected, rect=None):
    page = doc.new_page(width=backgrounds.A4_WIDTH, height=backgrounds.A4_HEIGHT)
    page.insert_image(rect or page.rect, stream=selected.data, overlay=False)
    return page


@pytest.mark.parametrize("kind", ["acreditado", "no_acreditado"])
def test_selected_background_fills_every_page_and_preserves_content(synthetic_backgrounds, kind):
    selected = backgrounds.load_letterhead(kind)
    with pymupdf.open() as doc:
        for _ in range(3):
            page = doc.new_page(width=backgrounds.A4_WIDTH, height=backgrounds.A4_HEIGHT)
            page.insert_text((80, 180), "FICTICIO - SIN VALIDEZ")
        before = [page.get_text("words") for page in doc]
        backgrounds.apply_letterhead(doc, selected)
        assert backgrounds.verify_letterhead(doc, selected) == [1, 2, 3]
        assert [page.get_text("words") for page in doc] == before
        assert len({image["xref"] for page in doc for image in page.get_image_info(xrefs=True)}) == 1


def test_known_shifted_previous_background_is_replaced_without_doubling(synthetic_backgrounds):
    old = backgrounds.load_letterhead("acreditado")
    chosen = backgrounds.load_letterhead("no_acreditado")
    with pymupdf.open() as doc:
        page_with_background(doc, old, pymupdf.Rect(15, 25, 580, 817))
        backgrounds.apply_letterhead(doc, chosen)
        assert backgrounds.verify_letterhead(doc, chosen) == [1]
        images = doc[0].get_image_info(hashes=True)
        assert all(image["digest"] != old.digest for image in images)


def test_missing_background_on_second_page_is_rejected(synthetic_backgrounds):
    chosen = backgrounds.load_letterhead("acreditado")
    with pymupdf.open() as doc:
        page_with_background(doc, chosen)
        doc.new_page(width=backgrounds.A4_WIDTH, height=backgrounds.A4_HEIGHT)
        with pytest.raises(CertificateError, match="página 2"):
            backgrounds.verify_letterhead(doc, chosen)


@pytest.mark.parametrize("failure", ["wrong_type", "shifted", "duplicated", "covered", "reduced_resolution"])
def test_background_verification_rejects_incorrect_output(synthetic_backgrounds, failure):
    chosen = backgrounds.load_letterhead("acreditado")
    with pymupdf.open() as doc:
        other = backgrounds.load_letterhead("no_acreditado") if failure == "wrong_type" else chosen
        rect = pymupdf.Rect(10, 20, 585, 821) if failure == "shifted" else None
        page = page_with_background(doc, other, rect)
        if failure == "duplicated":
            page.insert_image(page.rect, stream=chosen.data)
        if failure == "covered":
            page.draw_rect(page.rect, fill=(1, 1, 1), overlay=True)
        if failure == "reduced_resolution":
            chosen = replace(chosen, width=chosen.width * 2)
        with pytest.raises(CertificateError):
            backgrounds.verify_letterhead(doc, chosen)


def test_wrong_page_orientation_is_rejected(synthetic_backgrounds):
    chosen = backgrounds.load_letterhead("acreditado")
    with pymupdf.open() as doc:
        doc.new_page(width=backgrounds.A4_HEIGHT, height=backgrounds.A4_WIDTH)
        with pytest.raises(CertificateError, match="A4 verticales"):
            backgrounds.apply_letterhead(doc, chosen)


def test_corrupt_official_asset_is_rejected(synthetic_backgrounds):
    synthetic_backgrounds["acreditado"][1].write_bytes(b"damaged file")
    with pytest.raises(CertificateError, match="cambió o está dañado"):
        backgrounds.load_letterhead("acreditado")


def test_fallback_keeps_saved_print_margins_and_rejects_missing_source_image(synthetic_backgrounds):
    chosen = backgrounds.load_letterhead("acreditado")
    variants = [workbook_bytes()]
    def add_picture(parts):
        sheet = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        ET.SubElement(sheet, f"{{{S}}}picture", {"{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "missing-background"})
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
    margins = []
    for variant in variants:
        prepared, data, _, _, _ = _prepare_workbook(variant, False, chosen)
        assert data == chosen.data
        with zipfile.ZipFile(BytesIO(prepared)) as archive:
            sheet = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            assert sheet.find(f"{{{S}}}picture") is None
            margins.append(sheet.find(f"{{{S}}}pageMargins").attrib)
    assert margins[0]["left"] == "0.75" and margins[0]["top"] == "1.4"
    assert margins[0]["header"] == "0.5"
    with pytest.raises(CertificateError, match="Falta la imagen del fondo"):
        _prepare_workbook(rewrite(workbook_bytes(), add_picture), False, chosen)


def test_existing_header_prevents_adding_official_background(synthetic_backgrounds):
    chosen = backgrounds.load_letterhead("no_acreditado")
    def add_header_image(parts):
        sheet = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        header = sheet.find(f"{{{S}}}headerFooter")
        if header is None:
            header = ET.SubElement(sheet, f"{{{S}}}headerFooter")
        ET.SubElement(header, f"{{{S}}}oddHeader").text = "&CENCABEZADO FICTICIO"
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
    prepared, background, footer, _, _ = _prepare_workbook(rewrite(workbook_bytes(), add_header_image), False, chosen)
    assert background is None and footer == ""
    with zipfile.ZipFile(BytesIO(prepared)) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        assert sheet.find(f"{{{S}}}headerFooter/{{{S}}}oddHeader").text == "&CENCABEZADO FICTICIO"


@pytest.mark.parametrize("kind", ["acreditado", "no_acreditado"])
def test_excel_final_pdf_has_selected_background_and_first_page_qr(synthetic_backgrounds, kind):
    url = "https://drive.google.com/file/d/SYNTHETIC_SIN_VALIDEZ/view"
    result = prepare_document(workbook_bytes(), "xlsx", "", url, certificate_type=kind)
    assert result.certificate_type == kind
    assert result.background_pages == [1, 2]
    assert result.stamp_pages == [1]
    with pymupdf.open(stream=result.data, filetype="pdf") as doc:
        assert backgrounds.verify_letterhead(doc, backgrounds.load_letterhead(kind)) == [1, 2]
        assert "17,5" in doc[0].get_text()
        decoded = []
        for index, page in enumerate(doc):
            image = Image.open(BytesIO(page.get_pixmap(dpi=150).tobytes("png")))
            decoded.extend((index + 1, barcode.text) for barcode in zxingcpp.read_barcodes(image))
        assert decoded == [(1, url)]


@pytest.mark.parametrize("kind", ["", "unknown", "../assets/fondo-acreditado.jpeg"])
def test_excel_type_is_required_before_google_requests(kind):
    drive = FakeDrive()
    with TestClient(create_app(certificate_storage=drive.storage())) as client:
        response = client.post("/api/certificates", files={"excel": ("test.xlsx", workbook_bytes())},
                               data={"tipoCertificado": kind})
        assert response.status_code == 422
        assert "Selecciona" in response.json()["detail"]
        assert drive.token_requests == 0
        assert list(drive.files) == [ROOT]


def test_invalid_final_background_never_publishes_a_drive_record(monkeypatch):
    with pymupdf.open() as doc:
        doc.new_page(width=backgrounds.A4_WIDTH, height=backgrounds.A4_HEIGHT).insert_text((200, 300), "Escanee este QR")
        missing_background_pdf = doc.tobytes()
    # Simulate a converter regression; the final independent check must block it.
    monkeypatch.setattr("app.documents.convert_excel", lambda *args, **kwargs: ConvertedExcel(missing_background_pdf, False, []))
    drive = FakeDrive()
    with TestClient(create_app(certificate_storage=drive.storage())) as client:
        client.app.state.executor.shutdown(wait=True)
        client.app.state.executor = ThreadPoolExecutor(max_workers=1)
        response = client.post("/api/certificates", files={"excel": ("test.xlsx", workbook_bytes())},
                               data={"tipoCertificado": "acreditado"})
        assert response.status_code == 422
        assert "fondo oficial" in response.json()["detail"]
    assert list(drive.files) == [ROOT]
    assert drive.permissions == []
