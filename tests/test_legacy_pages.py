"""Legacy manual page breaks must preserve artwork and complete text."""
from io import BytesIO
import xml.etree.ElementTree as ET

from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Font, Alignment, Border, Side
from PIL import Image
import pymupdf
import pytest

from app.documents import prepare_document
from app.excel import NS, S, R, P, _prepare_workbook, inspect_excel
from app.excel_composition import fit_manual_numbered_pages, validate_table_text, validate_saved_labels
from app.pdf import CertificateError, prepare_pdf
from test_excel import rewrite, image_bytes
from test_excel_composition import numbered_certificate


def legacy_workbook(*, crossing=False, wrong_label=False):
    book = load_workbook(BytesIO(numbered_certificate()))
    sheet = book.active
    sheet.print_title_rows = None
    sheet['D51'] = None
    sheet.row_dimensions[1].height = 65
    sheet.row_breaks.append(Break(id=59))
    sheet.row_breaks.append(Break(id=117))
    sheet.merge_cells('A150:D150')
    sheet['A150'] = 'Máximo Error Absoluto de Histéresis'
    sheet['A150'].font = Font(name='Arial',size=10,bold=True)
    sheet['A150'].alignment = Alignment(horizontal='left')
    sheet['A150'].border = Border(left=Side(style='thin'),right=Side(style='thin'),
                                  top=Side(style='thin'),bottom=Side(style='thin'))
    # Four bright corners make a clipped signature surrogate detectable visually.
    image = Image.new('RGB', (200, 120), 'blue')
    image.paste('red', (0, 105, 200, 120))
    stream = BytesIO();image.save(stream, 'PNG')
    picture = ExcelImage(BytesIO(stream.getvalue()))
    picture.anchor = TwoCellAnchor(_from=AnchorMarker(col=4, row=47),
                                  to=AnchorMarker(col=9, row=60 if crossing else 58))
    sheet.add_image(picture)
    if wrong_label:
        sheet['H65'] = 'Pagina 3 de 3'
        sheet['H118'] = 'Pagina 2 de 3'
    output = BytesIO();book.save(output)
    def header(parts):
        workbook = ET.fromstring(parts['xl/workbook.xml'])
        names = workbook.find('s:definedNames',NS)
        for name in list(names):
            if name.get('name')=='_xlnm.Print_Titles':
                names.remove(name)
        parts['xl/workbook.xml'] = ET.tostring(workbook)
        root = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        hf = ET.SubElement(root, f'{{{S}}}headerFooter')
        ET.SubElement(hf, f'{{{S}}}oddHeader').text = '&C&G'
        ET.SubElement(root, f'{{{S}}}legacyDrawingHF', {f'{{{R}}}id':'HF'})
        parts['xl/worksheets/sheet1.xml'] = ET.tostring(root)
        rels = ET.fromstring(parts['xl/worksheets/_rels/sheet1.xml.rels'])
        ET.SubElement(rels, f'{{{P}}}Relationship', Id='HF', Type=R+'/vmlDrawing', Target='../drawings/header.vml')
        parts['xl/worksheets/_rels/sheet1.xml.rels'] = ET.tostring(rels)
        parts['xl/drawings/header.vml'] = b'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape id="CH" style="width:595pt;height:842pt"><v:imagedata o:relid="im0"/></v:shape></xml>'
        parts['xl/drawings/_rels/header.vml.rels'] = f'<Relationships xmlns="{P}"><Relationship Id="im0" Type="{R}/image" Target="../media/header.png"/></Relationships>'.encode()
        parts['xl/media/header.png'] = image_bytes((248,351),'#eef3f0')
        types = ET.fromstring(parts['[Content_Types].xml'])
        ET.SubElement(types,'{http://schemas.openxmlformats.org/package/2006/content-types}Default',Extension='vml',ContentType='application/vnd.openxmlformats-officedocument.vmlDrawing')
        parts['[Content_Types].xml'] = ET.tostring(types)
    return rewrite(output.getvalue(), header)


def test_manual_pages_recover_spilled_signature_and_add_qr_without_an_extra_page():
    url = 'https://drive.google.com/file/d/FICTITIOUS_LEGACY_REVIEW/view'
    result = prepare_document(legacy_workbook(), 'xlsx', '', url)
    assert any('saltos guardados' in warning for warning in result.warnings)
    assert any('no incluía QR' in warning for warning in result.warnings)
    with pymupdf.open(stream=result.data, filetype='pdf') as doc:
        assert len(doc) == 3
        for i,page in enumerate(doc):
            assert f'Pagina {i+1} de 3' in page.get_text()
        assert 'PRUEBA IBP FICTICIA' in doc[1].get_text()
        assert 'OBSERVACIONES FICTICIAS' in doc[2].get_text()
        assert 'FIN DEL DOCUMENTO FICTICIO' in doc[2].get_text()
        assert 'Máximo Error Absoluto de Histéresis' in doc[2].get_text()
        signature = next(info for info in doc[0].get_image_info() if (info['width'],info['height']) == (200,120))
        area = pymupdf.Rect(signature['bbox'])
        assert doc[0].rect.contains(area)
        bottom = pymupdf.Rect(area.x0,area.y1-5,area.x1,area.y1-1)
        pix = doc[0].get_pixmap(clip=bottom)
        pixels = Image.frombytes('RGB',(pix.width,pix.height),pix.samples)
        assert any(r>200 and g<50 and b<50 for r,g,b in
                   (pixels.getpixel((x,y)) for y in range(pixels.height) for x in range(pixels.width)))
        links = [link for page in doc for link in page.get_links() if link.get('uri')==url]
        assert len(links)==1
        qr = pymupdf.Rect(links[0]['from'])
        assert not (qr+(-4,-4,4,4)).intersects(area)


@pytest.mark.parametrize('options', [{'crossing':True},{'wrong_label':True}])
def test_conflicting_blocks_are_not_forced_into_three_pages(options):
    data = legacy_workbook(**options)
    layout = inspect_excel(data)
    from dataclasses import replace
    layout = replace(layout,page_numbers=(1,2,3))
    prepared,*_ = _prepare_workbook(data,False)
    assert fit_manual_numbered_pages(prepared, layout) is None


def test_trailing_spaces_do_not_reject_letters_inside_a_table():
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.draw_rect(pymupdf.Rect(40,100,150,125))
        page.draw_rect(pymupdf.Rect(150,100,260,125))
        page.insert_text((45,117),'Complete label'+' '*40,fontsize=10)
        validate_table_text(doc)


def test_repaired_labels_check_every_occurrence_and_the_last_letter():
    book = Workbook()
    sheet = book.active
    sheet.title = 'Certificado'
    label = 'Máximo Error Absoluto de Histéresis'
    sheet['A1'] = sheet['A2'] = label
    sheet.print_area = 'A1:F2'
    data = BytesIO();book.save(data)
    layout = inspect_excel(data.getvalue())
    prepared,*_ = _prepare_workbook(data.getvalue(),False)
    for last in [label,label[:-1]]:
        with pymupdf.open() as doc:
            page = doc.new_page()
            page.insert_text((50,100),label)
            page.insert_text((50,130),last)
            if last==label:
                validate_saved_labels(doc,prepared,layout)
            else:
                with pytest.raises(CertificateError,match='texto de la plantilla'):
                    validate_saved_labels(doc,prepared,layout)


def test_automatic_qr_preserves_artwork_and_rejects_an_occupied_page():
    with pymupdf.open() as doc:
        page = doc.new_page()
        bg = image_bytes((248,351),'#e8eee9')
        page.insert_image(page.rect,stream=bg,overlay=False)
        digest = page.get_image_info(hashes=True)[0]['digest']
        page.insert_text((50,120),'FICTICIO SIN VALIDEZ QCP-2026-0000')
        # A real image covering the body must never be mistaken for stationery.
        page.insert_image(pymupdf.Rect(40,150,560,800),stream=image_bytes((200,250),'blue'))
        with pytest.raises(CertificateError,match='espacio libre'):
            prepare_pdf(doc.tobytes(),'','https://example.com/test',template=True,
                        background_digests=(digest,),automatic_qr=True)
