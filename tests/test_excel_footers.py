from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
import pymupdf
import pytest

from app.documents import prepare_document
from app.excel import NS,S,_prepare_workbook,_read_parts
from app.excel_footers import (FormatFooter, take_format_footers, general_data_footer,
                               verify_format_footers, entered_footer)
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


def test_data_sheet_is_read_without_using_the_registers_own_form_identifier():
    data=workbook_bytes()
    def add(parts):
        book=ET.fromstring(parts['xl/workbook.xml'])
        entry=book.find('s:sheets/s:sheet',NS);entry.set('name','Datos generales');entry.set('state','hidden')
        parts['xl/workbook.xml']=ET.tostring(book)
        sheet=ET.fromstring(parts['xl/worksheets/sheet1.xml']);cells=sheet.find('s:sheetData',NS)
        for rownum,entries in [(4,[('A4','Código del formato del certificado:'),('B4','F-GCA-10.99')]),(5,[('A5','Versión:'),('B5','07')])]:
            row=ET.SubElement(cells,f'{{{S}}}row',r=str(rownum))
            for ref,text in entries:
                cell=ET.SubElement(row,f'{{{S}}}c',r=ref,t='inlineStr');ET.SubElement(ET.SubElement(cell,f'{{{S}}}is'),f'{{{S}}}t').text=text
        parts['xl/worksheets/sheet1.xml']=ET.tostring(sheet)
    source=rewrite(data,add)
    assert general_data_footer(_read_parts(source))==FormatFooter('F-GCA-10.99','07')
    bad=rewrite(source,lambda p:p.update({'xl/worksheets/sheet1.xml':p['xl/worksheets/sheet1.xml'].replace(b'F-GCA-10.99',b'F-GCA-07.02')}))
    assert general_data_footer(_read_parts(bad)) is None


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


def test_no_universal_code_or_version_is_invented():
    assert entered_footer('','') is None
    assert general_data_footer(_read_parts(workbook_bytes())) is None
