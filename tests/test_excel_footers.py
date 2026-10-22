from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
import pymupdf
import pytest

from app.documents import prepare_document
from app.excel import NS,S,_prepare_workbook,_read_parts
from app.excel_footers import (FormatFooter, take_format_footers, general_data_footer,
                               apply_format_footers, verify_format_footers, entered_footer)
from app.main import create_app
from app.pdf import CertificateError
from test_excel import workbook_bytes,rewrite


def footer_book():
    def add(parts):
        sheet=ET.fromstring(parts['xl/worksheets/sheet2.xml'])
        hf=ET.SubElement(sheet,f'{{{S}}}headerFooter')
        ET.SubElement(hf,f'{{{S}}}oddFooter').text='&LAVISO ORIGINAL&RCódigo: F-GCA-10.04\nVersión: 00\nPágina &P de &N'
        parts['xl/worksheets/sheet2.xml']=ET.tostring(sheet)
    return rewrite(workbook_bytes(),add)


def test_native_control_footer_is_kept_below_the_body_on_every_final_page():
    data=footer_book()
    result=prepare_document(data,'xlsx','','https://example.com/FOOTER',format_code='F-GCA-10.88',format_version='99')
    with pymupdf.open(stream=result.data,filetype='pdf') as doc:
        assert len(doc)==2
        verify_format_footers(doc,[FormatFooter('F-GCA-10.04','00')]*2)
        assert all('F-GCA-10.88' not in p.get_text() for p in doc)
        assert all('AVISO ORIGINAL' in p.get_text() for p in doc)
    prepared,*_=_prepare_workbook(data,False)
    with zipfile.ZipFile(BytesIO(prepared)) as z:
        sheet=ET.fromstring(z.read('xl/worksheets/sheet2.xml'))
        hf=sheet.find('s:headerFooter/s:oddFooter',NS).text
        assert 'AVISO ORIGINAL' in hf and 'F-GCA-10.04' not in hf and '&P' not in hf


def test_explicit_certificate_metadata_takes_priority_over_a_generic_register_heading():
    data=workbook_bytes()
    def add(parts):
        book=ET.fromstring(parts['xl/workbook.xml'])
        entry=book.find('s:sheets/s:sheet',NS);entry.set('name','Datos generales');entry.set('state','hidden')
        parts['xl/workbook.xml']=ET.tostring(book)
        sheet=ET.fromstring(parts['xl/worksheets/sheet1.xml']);cells=sheet.find('s:sheetData',NS)
        for rownum,entries in [(1,[('T1','CODIGO: F-GCA-07-04')]),(2,[('T2','VERSION: 01')]),(4,[('A4','Código del formato del certificado:'),('B4','F-GCA-10.99')]),(5,[('A5','Versión:'),('B5','07')])]:
            row=ET.SubElement(cells,f'{{{S}}}row',r=str(rownum))
            for ref,text in entries:
                cell=ET.SubElement(row,f'{{{S}}}c',r=ref,t='inlineStr');ET.SubElement(ET.SubElement(cell,f'{{{S}}}is'),f'{{{S}}}t').text=text
        parts['xl/worksheets/sheet1.xml']=ET.tostring(sheet)
    source=rewrite(data,add)
    assert general_data_footer(_read_parts(source))==FormatFooter('F-GCA-10.99','07')
    alternate=rewrite(source,lambda p:p.update({'xl/worksheets/sheet1.xml':p['xl/worksheets/sheet1.xml'].replace(b'F-GCA-10.99',b'F-GCA-07.02')}))
    assert general_data_footer(_read_parts(alternate))==FormatFooter('F-GCA-07.02','07')


def test_missing_format_fields_continue_with_page_numbers_and_explicitly_supplied_fields(tmp_path):
    with TestClient(create_app(tmp_path,'https://example.com')) as client:
        r=client.post('/api/certificates',files={'excel':('test.xlsx',workbook_bytes())},
                      data={'tipoCertificado':'acreditado','codigoFormato':'F-GCA-10.88','versionFormato':'03'})
        assert r.status_code==201,r.text
        with pymupdf.open(stream=client.get(r.json()['downloadUrl']).content,filetype='pdf') as doc:
            verify_format_footers(doc,[FormatFooter('F-GCA-10.88','03')]*len(doc))
        invalid=client.post('/api/certificates',files={'excel':('test.xlsx',workbook_bytes())},data={'codigoFormato':'F-GCA-10.88'})
        assert invalid.status_code==422 and 'código y la versión' in invalid.json()['detail']


def test_footer_verification_rejects_missing_or_incorrect_numbering():
    with pymupdf.open() as doc:
        doc.new_page()
        with pytest.raises(CertificateError,match='verificar'):
            verify_format_footers(doc,[FormatFooter('F-GCA-10.04','00')])


@pytest.mark.parametrize('footer', [FormatFooter('F-GCA-10.10','00'), None])
def test_control_footer_is_a_compact_aligned_block_on_every_page(footer):
    with pymupdf.open() as doc:
        for width,height in [(595,842),(842,595),(595,842),(595,842)]:
            page = doc.new_page(width=width,height=height)
            page.draw_line((20,height-27),(width-20,height-27),width=.4)
        apply_format_footers(doc,[footer]*len(doc))
        for index,page in enumerate(doc):
            labels = [f'Código: {footer.code}',f'Versión: {footer.version}'] if footer else []
            labels.append(f'Página {index+1} de {len(doc)}')
            boxes = [page.search_for(label)[0] for label in labels]
            assert all(box.x1 == pytest.approx(page.rect.width-7,abs=.01) for box in boxes)
            assert all(page.rect.contains(box) and box.y0 >= page.rect.height-44 for box in boxes)
            assert page.rect.height-25 < boxes[-1].y1 < page.rect.height-15
            rule = pymupdf.Rect(20,page.rect.height-27.2,page.rect.width-20,page.rect.height-26.8)
            assert not any(box.intersects(rule) for box in boxes)
            pitches = [lower.y0-upper.y0 for upper,lower in zip(boxes,boxes[1:])]
            assert all(7 <= pitch <= 10 for pitch in pitches)
            assert len(set(round(pitch,2) for pitch in pitches)) <= 1
            for upper,lower in zip(boxes,boxes[1:]):
                assert upper.y1 < lower.y0


def test_no_universal_code_or_version_is_invented():
    assert entered_footer('','') is None
    assert general_data_footer(_read_parts(workbook_bytes())) is None


def test_hyphenated_format_code_is_kept_complete_and_other_footer_text_survives():
    sheet = ET.fromstring(f'<worksheet xmlns="{S}"><headerFooter><oddFooter>&amp;LAVISO ORIGINAL&amp;RCódigo: F-GCA-07-04\nVersión: 01\nPágina &amp;P de &amp;N</oddFooter></headerFooter></worksheet>')
    assert take_format_footers(sheet)=={'odd':FormatFooter('F-GCA-07-04','01')}
    remaining = sheet.find('s:headerFooter/s:oddFooter',NS).text
    assert 'AVISO ORIGINAL' in remaining
    assert 'F-GCA' not in remaining and '-04' not in remaining and '&P' not in remaining


@pytest.mark.parametrize('wrong_field',['LP-ALT0125-PC001','QCP-2026-0001','MAN-67'])
def test_footer_does_not_accept_certificate_numbers_or_equipment_codes(wrong_field):
    with pytest.raises(CertificateError,match='código y la versión del formato'):
        entered_footer(wrong_field,'01')
