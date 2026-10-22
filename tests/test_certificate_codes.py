"""Fictitious identifiers and cached formulas; no client workbooks."""

import xml.etree.ElementTree as ET

from fastapi.testclient import TestClient
import pymupdf
import pytest

from app.excel import NS, S
from app.excel_footers import FormatFooter, inspect_format_footer, verify_format_footers
from app.main import create_app
from app.pdf import CODE, CODE_PATTERN, MAX_CODE_LENGTH, CertificateError, normalized_code, prepare_pdf
from test_excel import workbook_bytes, rewrite
from test_portal import pdf_bytes


@pytest.mark.parametrize('old,new', [
    ('LP-ALT0125-PC001','LP-COMP0125-PC002'),
    ('LP-ALT0125-PC001','QCP-2026-0002'),
    ('QCP-2026-0001','LP-ALT0125-PC002'),
])
def test_both_code_families_are_detected_and_corrected_on_every_page(old,new):
    result = prepare_pdf(pdf_bytes((old,old)),new.lower(),'https://example.com/test')
    assert result.detected_codes == [old]
    assert result.final_code == new
    assert result.replacements == 2
    with pymupdf.open(stream=result.data,filetype='pdf') as doc:
        for page in doc:
            assert new in page.get_text() and old not in page.get_text()
            assert 'Contenido que debe conservarse' in page.get_text()


@pytest.mark.parametrize('invalid', [
    'LP-ALT0125-PC', 'LP-ALT0125-PC001extra', 'LP-ALT0125-PC001-otro',
    'LP-ALT0125-PC٠٠١', 'QCP-2026-00010', 'XQCP-2026-0001',
])
def test_incomplete_extended_or_non_ascii_identifiers_are_not_accepted(invalid):
    with pytest.raises(CertificateError,match='formato'):
        normalized_code(invalid)
    assert not CODE.search(invalid)


def test_spaced_legacy_code_is_recognized_without_matching_instrument_or_format_ids():
    text = 'LP- ALT0125- PC001 / QCP-P-001 / F-GCA-07-04'
    assert CODE.findall(text) == ['LP- ALT0125- PC001']


def registered_workbook(format_code='F-GCA-07-04',version='01'):
    def add_register(parts):
        register = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        row = ET.SubElement(register.find('s:sheetData',NS),f'{{{S}}}row',r='7')
        for ref,text in [('B7','Numero de Certificado:'),('E7','LP-ALT0125-PC001')]:
            cell = ET.SubElement(row,f'{{{S}}}c',r=ref,t='inlineStr')
            ET.SubElement(ET.SubElement(cell,f'{{{S}}}is'),f'{{{S}}}t').text = text
        for number,ref,text in [(1,'T1',f'CODIGO: {format_code}'),(2,'T2',f'VERSION: {version}')]:
            row = register.find(f"s:sheetData/s:row[@r='{number}']",NS)
            cell = ET.SubElement(row,f'{{{S}}}c',r=ref,t='inlineStr')
            ET.SubElement(ET.SubElement(cell,f'{{{S}}}is'),f'{{{S}}}t').text = text
        parts['xl/worksheets/sheet1.xml'] = ET.tostring(register)
        certificate = ET.fromstring(parts['xl/worksheets/sheet2.xml'])
        certificate.find("s:sheetData/s:row/s:c[@r='H3']/s:f",NS).text = 'Registro!E7'
        parts['xl/worksheets/sheet2.xml'] = ET.tostring(certificate)
    return rewrite(workbook_bytes('LP-ALT0125-PC001'),add_register)


@pytest.mark.parametrize('code,version',[('F-GCA-07-04','01'),('F-GCA-14.28','07')])
def test_register_format_code_and_version_are_read_without_fixed_defaults(code,version):
    assert inspect_format_footer(registered_workbook(code,version)) == FormatFooter(code,version)


def test_conflicting_format_metadata_requires_explicit_resolution():
    def conflicting(parts):
        register = ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        for number,text in [(10,'CODIGO: F-GCA-14.28'),(11,'VERSION: 07')]:
            row = ET.SubElement(register.find('s:sheetData',NS),f'{{{S}}}row',r=str(number))
            cell = ET.SubElement(row,f'{{{S}}}c',r=f'T{number}',t='inlineStr')
            ET.SubElement(ET.SubElement(cell,f'{{{S}}}is'),f'{{{S}}}t').text = text
        parts['xl/worksheets/sheet1.xml'] = ET.tostring(register)
    assert inspect_format_footer(rewrite(registered_workbook(),conflicting)) is None


def test_cached_register_code_and_manual_correction_use_the_cached_upload(tmp_path):
    with TestClient(create_app(tmp_path,'https://example.com')) as client:
        config = client.get('/api/config').json()
        assert config['certificateCodePattern'] == CODE_PATTERN
        assert config['maxCodeLength'] == MAX_CODE_LENGTH >= len('LP-COMP0125-PC002')
        for requested,expected in [('', 'LP-ALT0125-PC001'),('lp-comp0125-pc002','LP-COMP0125-PC002')]:
            workbook = registered_workbook(format_code='' if requested else 'F-GCA-07-04')
            inspected = client.post('/api/excel/inspect',files={'excel':('otro-nombre.xlsx',workbook)})
            assert inspected.status_code == 200
            if requested:
                assert 'formatFooter' not in inspected.json()
            else:
                assert inspected.json()['formatFooter']=={'code':'F-GCA-07-04','version':'01'}
            response = client.post('/api/certificates',data={
                'excelToken':inspected.json()['uploadToken'],
                'tipoCertificado':'no_acreditado','codigoCorrecto':requested,
                **({'codigoFormato':'F-GCA-07-04','versionFormato':'01'} if requested else {}),
            })
            assert response.status_code == 201,response.text
            info = response.json()
            assert info['finalCode'] == expected
            assert info['detectedCodes'] == ['LP-ALT0125-PC001']
            assert info['replacements'] == (2 if requested else 0)
            with pymupdf.open(stream=client.get(info['downloadUrl']).content,filetype='pdf') as doc:
                assert len(doc) == 2
                verify_format_footers(doc,[FormatFooter('F-GCA-07-04','01')]*2)
                for page in doc:
                    assert expected in page.get_text()
                    assert 'OTHER SHEET' not in page.get_text()
                assert '17,5' in doc[0].get_text()


def test_unknown_codes_are_never_claimed_as_a_manual_replacement():
    result = prepare_pdf(pdf_bytes(('SIN CODIGO',)),'LP-ALT0125-PC001','https://example.com/test')
    assert result.final_code is None and result.replacements == 0
    assert result.warnings
