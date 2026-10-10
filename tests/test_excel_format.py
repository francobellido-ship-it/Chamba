"""Synthetic existing stationery and header pictures must survive conversion."""

from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
import pymupdf
import pytest

from app.backgrounds import load_letterhead
from app.excel import S, R, P, _prepare_workbook, inspect_excel
from app.excel_headers import header_pictures, restore_header_pictures
from app.main import create_app
from app.pdf import CertificateError, prepare_pdf
from test_excel import workbook_bytes, rewrite, image_bytes, qr_images


def header_workbook():
    def add(parts):
        sheet = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        header = ET.SubElement(sheet, f"{{{S}}}headerFooter", differentFirst="1", differentOddEven="1", scaleWithDoc="0")
        for tag, text in (("firstHeader", "&L&G&CENCABEZADO PRIMERA"), ("evenHeader", "&L&G&CENCABEZADO SEGUNDA"),
                          ("firstFooter", "&LPIE PRIMERA&R&G"), ("evenFooter", "&LPIE SEGUNDA&R&G")):
            ET.SubElement(header, f"{{{S}}}{tag}").text = text
        ET.SubElement(sheet, f"{{{S}}}legacyDrawingHF", {f"{{{R}}}id": "HF"})
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
        rels = ET.fromstring(parts["xl/worksheets/_rels/sheet2.xml.rels"])
        ET.SubElement(rels, f"{{{P}}}Relationship", Id="HF", Type=R + "/vmlDrawing", Target="../drawings/header.vml")
        parts["xl/worksheets/_rels/sheet2.xml.rels"] = ET.tostring(rels)
        shapes, images = [], []
        for index, (identity, color) in enumerate((("LHF", "yellow"), ("LHE", "orange"), ("RFF", "cyan"), ("RFE", "purple"))):
            shapes.append(f'<v:shape id="{identity}" style="width:90pt;height:37.5pt;margin-left:0;margin-top:0"><v:imagedata o:relid="im{index}"/></v:shape>')
            images.append(f'<Relationship Id="im{index}" Type="{R}/image" Target="../media/header{index}.png"/>')
            parts[f"xl/media/header{index}.png"] = image_bytes((120, 50), color)
        parts["xl/drawings/header.vml"] = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' + ''.join(shapes) + '</xml>').encode()
        parts["xl/drawings/_rels/header.vml.rels"] = (f'<Relationships xmlns="{P}">' + ''.join(images) + '</Relationships>').encode()
        content = ET.fromstring(parts["[Content_Types].xml"])
        ET.SubElement(content, "{http://schemas.openxmlformats.org/package/2006/content-types}Default", Extension="vml", ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing")
        parts["[Content_Types].xml"] = ET.tostring(content)
    return rewrite(workbook_bytes(), add)


def test_vml_variants_position_resolution_and_no_duplication():
    data = header_workbook()
    settings = header_pictures(data)
    assert inspect_excel(data).preserve_format
    with pymupdf.open() as doc:
        doc.new_page()
        doc.new_page()
        restore_header_pictures(doc, settings)
        restore_header_pictures(doc, settings)
        for index, page in enumerate(doc):
            pictures = page.get_image_info(hashes=True)
            assert len(pictures) == 2
            assert {image["digest"] for image in pictures} == {picture.digest for picture in settings[0] if picture.variant == ("first" if index == 0 else "even")}
            assert all((image["width"], image["height"]) == (120, 50) for image in pictures)
            areas = [pymupdf.Rect(image["bbox"]) for image in pictures]
            assert any(abs(area.x0 - 54) < .01 and abs(area.y0 - 36) < .01 for area in areas)


def test_native_existing_header_footer_images_do_not_require_background_choice(tmp_path):
    data = header_workbook()
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        inspection = client.post("/api/excel/inspect", files={"excel": ("test.xlsx", data)}).json()
        assert inspection["preserveFormat"] and not inspection["needsBackground"]
        response = client.post("/api/certificates", files={"excel": ("test.xlsx", data)})
        assert response.status_code == 201, response.text
        info = response.json()
        assert info["formatPreserved"] and not info["backgroundVerified"]
        assert info["certificateType"] is None and info["backgroundPages"] == []
        with pymupdf.open(stream=client.get(info["downloadUrl"]).content, filetype="pdf") as doc:
            assert len(doc) == 2
            assert "ENCABEZADO PRIMERA" in doc[0].get_text()
            assert "PIE PRIMERA" in doc[0].get_text()
            assert "ENCABEZADO SEGUNDA" in doc[1].get_text()
            assert "PIE SEGUNDA" in doc[1].get_text()
            for page in doc:
                assert sum((im["width"], im["height"]) == (120, 50) for im in page.get_image_info()) == 2
                assert not any(pymupdf.Rect(im["bbox"]).get_area() > page.rect.get_area() * .9 for im in page.get_image_info())
            assert [url for url, _ in qr_images(doc)] == [info["verificationUrl"]]


def test_existing_background_is_kept_even_with_opposite_selection():
    original = load_letterhead("acreditado")
    fallback = load_letterhead("no_acreditado")
    def add(parts):
        sheet = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        ET.SubElement(sheet, f"{{{S}}}picture", {f"{{{R}}}id": "background"})
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(sheet)
        rels = ET.fromstring(parts["xl/worksheets/_rels/sheet2.xml.rels"])
        ET.SubElement(rels, f"{{{P}}}Relationship", Id="background", Type=R + "/image", Target="../media/background.jpeg")
        parts["xl/worksheets/_rels/sheet2.xml.rels"] = ET.tostring(rels)
        parts["xl/media/background.jpeg"] = original.data
    prepared, background, _, _, _ = _prepare_workbook(rewrite(workbook_bytes(), add), False, fallback)
    assert background == original.data
    with zipfile.ZipFile(BytesIO(prepared)) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
        assert sheet.find(f"{{{S}}}pageMargins").get("top") == "1"


def test_large_unrelated_image_is_not_treated_as_stationery():
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_image(page.rect, stream=image_bytes((600, 850), "blue"), keep_proportion=False)
        page.insert_text((200, 300), "Escanee este QR")
        with pytest.raises(CertificateError, match="espacio libre"):
            prepare_pdf(doc.tobytes(), "", "https://example.com/test", template=True)
