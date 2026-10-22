"""Regression cases for static XLSM, hidden rows, OLE scope and page stationery."""

from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
import pymupdf
import pytest
import qrcode
import zxingcpp
from PIL import Image, ImageCms

from app.excel import NS, S, P, R, _prepare_workbook
from app.excel_headers import header_pictures, restore_header_pictures
from app.excel_pages import fit_stationery_content
from app.main import create_app
from app.pdf import CertificateError, prepare_pdf
from app.pdf_images import compact_backgrounds
from test_excel import workbook_bytes, rewrite, image_bytes
from test_excel_format import header_workbook


def test_small_printed_signature_keeps_its_full_native_pixel_resolution():
    from app.excel import convert_excel
    original=image_bytes((1000,450),'blue')
    data=rewrite(workbook_bytes(),lambda parts:parts.update({'xl/media/image2.png':original}))
    with pymupdf.open(stream=convert_excel(data).pdf,filetype='pdf')as doc:
        images=[image for page in doc for image in page.get_image_info(xrefs=True)
                if (image['width'],image['height'])==(1000,450)]
        assert len(images)==1
        actual=Image.open(BytesIO(doc.extract_image(images[0]['xref'])['image'])).convert('RGB')
        assert actual.tobytes()==Image.open(BytesIO(original)).convert('RGB').tobytes()


@pytest.mark.parametrize("hidden", [True, False])
def test_hidden_formula_errors_are_not_printed_and_visible_ones_name_the_cell(hidden):
    def change(parts):
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        row = root.find("s:sheetData/s:row[@r='5']", NS)
        if hidden:
            row.set("hidden", "1")
        cell = row.find("s:c[@r='D5']", NS)
        cell.set("t", "e")
        cell.find("s:v", NS).text = "#DIV/0!"
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    data = rewrite(workbook_bytes(), change)
    if hidden:
        prepared, *_ = _prepare_workbook(data, False)
        with zipfile.ZipFile(BytesIO(prepared)) as source:
            root = ET.fromstring(source.read("xl/worksheets/sheet2.xml"))
            assert root.find("s:sheetData/s:row[@r='5']", NS).get("hidden") == "1"
            assert root.find("s:sheetData/s:row[@r='5']/s:c[@r='D5']/s:f", NS) is None
    else:
        with pytest.raises(CertificateError, match=r"D5.*#DIV/0!"):
            _prepare_workbook(data, False)


@pytest.mark.parametrize("selected", [True, False])
def test_ole_outside_certificate_is_omitted_but_printed_ole_requires_an_image(selected):
    def change(parts):
        path = f"xl/worksheets/sheet{2 if selected else 1}.xml"
        root = ET.fromstring(parts[path])
        objects = ET.SubElement(root, f"{{{S}}}oleObjects")
        ET.SubElement(objects, f"{{{S}}}oleObject", progId="PBrush", shapeId="1")
        parts[path] = ET.tostring(root)
        parts["xl/embeddings/oleObject1.bin"] = b"never opened or executed"
    data = rewrite(workbook_bytes(), change)
    if selected:
        with pytest.raises(CertificateError, match="hoja Certificado contiene objetos incrustados"):
            _prepare_workbook(data, False)
    else:
        prepared, *_ = _prepare_workbook(data, False)
        with zipfile.ZipFile(BytesIO(prepared)) as source:
            assert not any(name.startswith("xl/embeddings/") for name in source.namelist())
            assert "xl/worksheets/sheet1.xml" not in source.namelist()
            assert "xl/worksheets/sheet2.xml" in source.namelist()


def test_plain_print_names_and_excessive_horizontal_page_count_are_normalized():
    def change(parts):
        parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b"_xlnm.Print_", b"Print_")
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        root.find("s:pageSetup", NS).set("fitToWidth", "9")
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    prepared, *_ = _prepare_workbook(rewrite(workbook_bytes(), change), False)
    with zipfile.ZipFile(BytesIO(prepared)) as source:
        root = ET.fromstring(source.read("xl/workbook.xml"))
        assert {name.get("name") for name in root.findall("s:definedNames/s:definedName", NS)} == {"_xlnm.Print_Area", "_xlnm.Print_Titles"}
        setup = ET.fromstring(source.read("xl/worksheets/sheet2.xml")).find("s:pageSetup", NS)
        assert setup.get("fitToWidth") == "1" and setup.get("fitToHeight") == "0"


def full_page_workbook(marker=True, section="CH"):
    def change(parts):
        parts["xl/drawings/header.vml"] = (f'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="{section}" style="width:638.5pt;height:972pt"><v:imagedata o:relid="im0"/></v:shape></xml>').encode()
        parts["xl/media/header0.png"] = image_bytes((248, 351), "#d8e8e1")
        root = ET.fromstring(parts["xl/worksheets/sheet2.xml"])
        header = root.find("s:headerFooter", NS)
        for node in list(header):
            header.remove(node)
        if marker:
            ET.SubElement(header, f"{{{S}}}oddHeader").text = "&C&G"
        header.set("differentFirst", "0")
        header.set("differentOddEven", "0")
        root.find("s:pageSetup", NS).set("scale", "88")
        parts["xl/worksheets/sheet2.xml"] = ET.tostring(root)
    return rewrite(header_workbook(), change)


@pytest.mark.parametrize("marker", [True, False])
@pytest.mark.parametrize("section", ["LH", "CH", "RH"])
def test_page_sized_header_fills_each_page_without_scaling_cropping_or_pixel_changes(marker, section):
    data = full_page_workbook(marker, section)
    settings = header_pictures(data)
    assert settings[0][0].full_page
    prepared, *_ = _prepare_workbook(data, False)
    with zipfile.ZipFile(BytesIO(prepared)) as source:
        root = ET.fromstring(source.read("xl/worksheets/sheet2.xml"))
        assert all("&G" not in (node.text or "") for node in root.find("s:headerFooter", NS))
        assert root.find("s:legacyDrawingHF", NS) is None
        assert "xl/media/header0.png" not in source.namelist()
    with pymupdf.open() as doc:
        for _ in range(2):
            page = doc.new_page(width=595.3, height=841.89)
            page.insert_text((50, 25), "CONTENIDO FICTICIO")
        fit_stationery_content(doc, settings)
        restore_header_pictures(doc, settings)
        restore_header_pictures(doc, settings)
        for page in doc:
            assert page.search_for("CONTENIDO FICTICIO")[0].y0 >= 100
            images = page.get_image_info(hashes=True)
            assert len(images) == 1
            assert images[0]["digest"] == settings[0][0].digest
            assert (images[0]["width"], images[0]["height"]) == (248, 351)
            assert max(abs(a-b) for a,b in zip(images[0]["bbox"],page.rect)) < .01


def test_added_letterhead_fits_content_after_export_without_losing_text_or_creating_pages():
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_text((60, 55), "INICIO FICTICIO")
        page.insert_text((60, 810), "FINAL FICTICIO")
        before = [word[4] for word in page.get_text("words")]
        fit_stationery_content(doc, ([], {}, False, False), added_letterhead=True)
        assert len(doc) == 1
        after = doc[0].get_text("words")
        assert [word[4] for word in after] == before
        assert min(word[1] for word in after) >= 100
        assert max(word[3] for word in after) <= doc[0].rect.height - 42


def test_existing_drive_qr_is_replaced_in_its_slot_without_a_text_guide():
    old = "https://drive.google.com/file/d/EXISTING_FICTITIOUS_FILE/view"
    new = "https://drive.google.com/file/d/NEW_FICTITIOUS_FILE/view"
    buffer = BytesIO()
    qrcode.make(old).save(buffer, format="PNG")
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_text((50, 70), "QCP- 2026-0001")
        page.insert_text((50, 100), "DATOS FICTICIOS QUE DEBEN CONSERVARSE")
        slot = pymupdf.Rect(300, 500, 374, 574)
        page.insert_image(slot, stream=buffer.getvalue())
        result = prepare_pdf(doc.tobytes(), "", new, template=True)
    with pymupdf.open(stream=result.data,filetype="pdf") as doc:
        assert result.final_code == "QCP-2026-0001"
        assert len(doc)==1 and result.stamp_pages==[1]
        assert "DATOS FICTICIOS QUE DEBEN CONSERVARSE" in doc[0].get_text()
        codes=zxingcpp.read_barcodes(Image.open(BytesIO(doc[0].get_pixmap(dpi=150).tobytes("png"))))
        assert [code.text for code in codes] == [new]
        assert any(link.get("uri")==new for link in doc[0].get_links())


def test_spaced_certificate_code_can_be_replaced_without_leaving_the_previous_code():
    with pymupdf.open() as doc:
        doc.new_page().insert_text((50,70),"QCP- 2026-0001")
        result=prepare_pdf(doc.tobytes(),"QCP-2026-0002","https://example.com/test")
    assert result.final_code=="QCP-2026-0002" and result.replacements==1
    with pymupdf.open(stream=result.data,filetype="pdf") as doc:
        assert "QCP-2026-0002" in doc[0].get_text() and "0001" not in doc[0].get_text()


def test_xlsm_static_print_upload_and_filename_discrepancy_are_visible(tmp_path):
    data = workbook_bytes("QCP-2026-0001")
    with TestClient(create_app(tmp_path,"https://example.com")) as client:
        inspection = client.post("/api/excel/inspect",files={"excel":("test.xlsm",data)})
        assert inspection.status_code==200
        response = client.post("/api/certificates",files={"excel":("QCP-2026-0002.xlsm",data)},data={"tipoCertificado":"no_acreditado"})
        assert response.status_code==201,response.text
        info=response.json()
        assert info["filename"]=="QCP-2026-0002.pdf" and info["finalCode"]=="QCP-2026-0001"
        assert any("nombre del archivo" in warning and "QCP-2026-0001" in warning for warning in info["warnings"])


def test_redundant_jpeg_icc_is_removed_without_changing_pixels_or_pdf_colour_space():
    image = Image.new("RGB", (248, 351), "#b6d7c5")
    profile = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    data = BytesIO()
    image.save(data, format="JPEG", icc_profile=profile)
    with pymupdf.open() as doc:
        page = doc.new_page()
        xref = page.insert_image(page.rect, stream=data.getvalue())
        digest = page.get_image_info(hashes=True)[0]["digest"]
        colorspace = doc.xref_get_key(xref, "ColorSpace")
        original = doc.xref_stream_raw(xref)
        before = page.get_pixmap().samples
        compact_backgrounds(doc, (digest,))
        assert len(doc.xref_stream_raw(xref)) < len(original)
        assert doc.xref_get_key(xref, "ColorSpace") == colorspace
        assert doc.xref_stream_raw(xref).split(b"\xff\xda", 1)[1] == original.split(b"\xff\xda", 1)[1]
        assert page.get_pixmap().samples == before
        assert page.get_image_info(hashes=True)[0]["digest"] == digest
        compact = doc.xref_stream_raw(xref)
        compact_backgrounds(doc, (digest,))
        assert doc.xref_stream_raw(xref) == compact


@pytest.mark.parametrize("scale,rows", [(84,35), (50,80)])
def test_zero_margin_background_keeps_right_columns_without_enlarging_saved_pages(scale, rows):
    from app.excel import convert_excel
    def wide(parts):
        sheet = ET.fromstring(parts['xl/worksheets/sheet2.xml'])
        sheet.find('s:pageMargins', NS).attrib.update(left='0', right='0', top='0', bottom='0')
        sheet.find('s:sheetPr/s:pageSetUpPr', NS).set('fitToPage', '0')
        sheet.find('s:pageSetup', NS).attrib.update(scale=str(scale), fitToHeight='0')
        for column in sheet.findall('s:cols/s:col', NS):
            column.set('width', '11')
        sheet.remove(sheet.find('s:rowBreaks', NS))
        row = sheet.find("s:sheetData/s:row[@r='25']", NS)
        if row is None:
            row = ET.SubElement(sheet.find('s:sheetData', NS), f'{{{S}}}row', r='25')
        cell = ET.SubElement(row, f'{{{S}}}c', r='K25', t='inlineStr')
        ET.SubElement(ET.SubElement(cell, f'{{{S}}}is'), f'{{{S}}}t').text = 'BORDE DERECHO'
        ET.SubElement(sheet, f'{{{S}}}picture', {f'{{{R}}}id': 'background'})
        parts['xl/worksheets/sheet2.xml'] = ET.tostring(sheet)
        book = ET.fromstring(parts['xl/workbook.xml'])
        book.find("s:definedNames/s:definedName[@name='_xlnm.Print_Area']", NS).text = f"'Certificado'!$A$1:$L${rows}"
        parts['xl/workbook.xml'] = ET.tostring(book)
        relationships = ET.fromstring(parts['xl/worksheets/_rels/sheet2.xml.rels'])
        ET.SubElement(relationships, f'{{{P}}}Relationship', Id='background', Type=R+'/image', Target='../media/background.png')
        parts['xl/worksheets/_rels/sheet2.xml.rels'] = ET.tostring(relationships)
        parts['xl/media/background.png'] = image_bytes((248,351), '#d8e8e1')
    data = rewrite(workbook_bytes(), wide)
    converted = convert_excel(data)
    with pymupdf.open(stream=converted.pdf,filetype='pdf') as doc:
        assert len(doc) == 1
        assert 'BORDE DERECHO' in doc[0].get_text()
        assert '17,5' in doc[0].get_text()
        words = doc[0].get_text('words')
        assert all(word[0] >= 0 and word[2] <= doc[0].rect.width for word in words)
        body_words = [word for word in words if word[1] < doc[0].rect.height - 44]
        assert min(word[1] for word in body_words) >= 100
        assert max(word[3] for word in body_words) <= doc[0].rect.height - 42
        assert 'Página 1 de 1' in doc[0].get_text(clip=pymupdf.Rect(0,doc[0].rect.height-44,doc[0].rect.width,doc[0].rect.height))
        image = next(image for image in doc[0].get_image_info() if (image['width'], image['height']) == (248,351))
        assert max(abs(a-b) for a,b in zip(image['bbox'], doc[0].rect)) < .01


def test_shared_jpeg_compaction_keeps_other_images_and_all_page_pixels():
    profile = ImageCms.ImageCmsProfile(ImageCms.createProfile('sRGB')).tobytes()
    def jpeg(color):
        buffer = BytesIO()
        Image.new('RGB', (248, 351), color).save(buffer, format='JPEG', icc_profile=profile)
        return buffer.getvalue()
    with pymupdf.open() as doc:
        first = doc.new_page()
        selected = first.insert_image(first.rect, stream=jpeg('#b6d7c5'))
        other = first.insert_image(pymupdf.Rect(40, 100, 80, 150), stream=jpeg('#445577'))
        for _ in range(2):
            page = doc.new_page()
            page.insert_image(page.rect, xref=selected)
        digest = pymupdf.Pixmap(doc, selected).digest
        streams = {xref: doc.xref_stream_raw(xref) for xref in (selected, other)}
        before = [page.get_pixmap().samples for page in doc]
        compact_backgrounds(doc, (digest,))
        assert len(doc.xref_stream_raw(selected)) < len(streams[selected])
        assert doc.xref_stream_raw(other) == streams[other]
        assert [page.get_pixmap().samples for page in doc] == before
        assert all(selected in [image[0] for image in page.get_images()] for page in doc)


def test_active_footer_picture_survives_removing_full_page_header_marker():
    def add_footer(parts):
        path = 'xl/drawings/header.vml'
        root = ET.fromstring(parts[path])
        shape = ET.SubElement(root, '{urn:schemas-microsoft-com:vml}shape',
                              id='RF', style='width:90pt;height:37.5pt')
        ET.SubElement(shape, '{urn:schemas-microsoft-com:vml}imagedata',
                      {'{urn:schemas-microsoft-com:office:office}relid': 'im1'})
        parts[path] = ET.tostring(root)
        path = 'xl/worksheets/sheet2.xml'
        sheet = ET.fromstring(parts[path])
        ET.SubElement(sheet.find('s:headerFooter', NS), f'{{{S}}}oddFooter').text = '&R&G'
        parts[path] = ET.tostring(sheet)
    data = rewrite(full_page_workbook(), add_footer)
    prepared, *_ = _prepare_workbook(data, False)
    with zipfile.ZipFile(BytesIO(prepared)) as archive:
        sheet = ET.fromstring(archive.read('xl/worksheets/sheet2.xml'))
        assert sheet.find('s:legacyDrawingHF', NS) is not None
        assert sheet.find('s:headerFooter/s:oddFooter', NS).text == '&R&G'
        assert 'xl/media/header1.png' in archive.namelist()
    assert len(header_pictures(prepared)[0]) == 2
