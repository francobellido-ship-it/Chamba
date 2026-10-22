"""Saved numeric powers and whole-cell superscripts in fictitious workbooks."""

from io import BytesIO
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.styles import Font,Alignment
import pymupdf

from app.excel import NS,_prepare_workbook,convert_excel
from test_excel import rewrite


def script_workbook():
    book=Workbook();sheet=book.active;sheet.title='Certificado'
    sheet['A1']='FICTICIO - SIN VALIDEZ';sheet['H2']='QCP-2026-0001'
    for column in 'ABCDEFGHIJKL':sheet.column_dimensions[column].width=8.3
    for row in range(1,35):sheet.row_dimensions[row].height=20
    for row,exponent in enumerate([2,-11,0,-3,12],5):
        sheet[f'A{row}']=f'POTENCIA {row-4}'
        sheet[f'H{row}']='10';sheet[f'H{row}'].font=Font(name='Arial',size=11)
        sheet[f'H{row}'].alignment=Alignment(horizontal='right')
        cell=sheet[f'I{row}'];cell.value=exponent
        cell.font=Font(name='Arial',size=11,vertAlign='superscript')
        cell.alignment=Alignment(horizontal='left')
    sheet['I9']='=6+6'  # A saved formula can itself supply an exponent.
    sheet['H10']='R';sheet['H10'].font=Font(name='Arial',size=11)
    sheet['H10'].alignment=Alignment(horizontal='right')
    sheet['I10']='n';sheet['I10'].font=Font(name='Arial',size=11,vertAlign='subscript')
    sheet['A12']='RESULTADO CUADRADO';sheet['D12']='=3^2';sheet['D12'].number_format='0.00'
    sheet['A13']='RESULTADO NEGATIVO';sheet['D13']='=10^-11';sheet['D13'].number_format='0.00000000000'
    sheet.merge_cells('D13:F13')
    sheet['D20']='Escanee este QR';sheet['D21']='Para consultar el certificado'
    sheet.print_area='A1:L32';sheet.page_setup.paperSize=sheet.PAPERSIZE_A4
    sheet.sheet_properties.pageSetUpPr.fitToPage=True
    sheet.page_setup.fitToWidth=1;sheet.page_setup.fitToHeight=1
    output=BytesIO();book.save(output)
    def saved(parts):
        root=ET.fromstring(parts['xl/worksheets/sheet1.xml'])
        for ref,value in [('D12','9'),('D13','1E-11'),('I9','12')]:
            root.find(f"s:sheetData/s:row/s:c[@r='{ref}']/s:v",NS).text=value
        parts['xl/worksheets/sheet1.xml']=ET.tostring(root)
    return rewrite(output.getvalue(),saved)


def test_whole_cell_superscripts_are_raised_and_saved_power_results_are_preserved():
    source=script_workbook()
    prepared,*_=_prepare_workbook(source,False)
    with ZipFile(BytesIO(prepared))as archive:
        root=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
    for row,exponent in enumerate([2,-11,0,-3,12],5):
        cell=root.find(f"s:sheetData/s:row/s:c[@r='I{row}']",NS)
        assert cell.get('t')=='inlineStr'
        assert cell.find('s:is/s:r/s:rPr/s:vertAlign',NS).get('val')=='superscript'
        assert cell.find('s:is/s:r/s:t',NS).text==str(exponent)
    for ref,value in [('D12','9'),('D13','1E-11')]:
        cell=root.find(f"s:sheetData/s:row/s:c[@r='{ref}']",NS)
        assert cell.find('s:f',NS)is None and cell.find('s:v',NS).text==value
    with pymupdf.open(stream=convert_excel(source).pdf,filetype='pdf')as doc:
        assert len(doc)==1
        text=doc[0].get_text()
        assert '9,00' in text and '0,00000000001' in text
        spans=[s for b in doc[0].get_text('dict')['blocks']for l in b.get('lines',[])for s in l['spans']]
        bases=sorted([s for s in spans if s['text']=='10'],key=lambda s:s['origin'][1])
        powers=sorted([s for s in spans if s['text']in {'2','-11','0','-3','12'}],key=lambda s:s['origin'][1])
        assert len(bases)==len(powers)==5
        for base,power in zip(bases,powers,strict=True):
            assert power['size']<base['size']*.8
            assert power['origin'][1]<base['origin'][1]-1
        base=next(s for s in spans if s['text']=='R')
        index=next(s for s in spans if s['text']=='n')
        assert index['size']<base['size']*.8
        assert index['origin'][1]>base['origin'][1]+.5
