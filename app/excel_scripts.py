"""Preserve whole-cell super/subscripts that Calc drops from numeric fonts."""

from copy import deepcopy
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import re
import xml.etree.ElementTree as ET

from app.excel_ranges import NS


def preserve_cell_scripts(sheet, styles, strings, value_of, printed, *, rich_strings=()):
    """Use rich text for saved script values; never evaluate an Excel formula."""
    fonts = styles.findall('s:fonts/s:font', NS)
    formats = styles.findall('s:cellXfs/s:xf', NS)
    numbers = {'0':'General', '1':'0', '2':'0.00'}
    numbers.update({item.get('numFmtId'):item.get('formatCode','')
                    for item in styles.findall('s:numFmts/s:numFmt',NS)})
    adjusted = []
    for cell in sheet.findall('s:sheetData/s:row/s:c', NS):
        if not printed(cell.get('r','')):
            continue
        style = formats[int(cell.get('s','0'))]
        font = fonts[int(style.get('fontId','0'))]
        script = font.find('s:vertAlign',NS)
        if script is None or script.get('val') not in {'superscript','subscript'}:
            continue
        text = value_of(cell,strings)
        if not text or cell.get('t') == 'e':
            continue
        if cell.get('t','n') == 'n':
            code = numbers.get(style.get('numFmtId','0'),'')
            code = re.sub(r'\[\$-[0-9A-Fa-f]+\]','',code)
            try:
                value = Decimal(text)
                if not value.is_finite() or abs(value.adjusted()) > 324 or len(code) > 255:
                    continue
                if code.casefold() == 'general':
                    text = format(value,'f')
                elif re.fullmatch(r'0+(?:\.0+)?',code):
                    decimals = len(code.partition('.')[2])
                    with localcontext() as context:
                        context.prec = max(50,len(value.as_tuple().digits)+abs(value.adjusted())+decimals+2)
                        text = format(value.quantize(Decimal(1).scaleb(-decimals),rounding=ROUND_HALF_UP),'f')
                    sign = '-' if text.startswith('-') else ''
                    text = sign+text.lstrip('-').zfill(len(code))
                    text = text.replace('.',',')
                else:
                    continue  # Do not guess date, unit or other custom formats.
            except InvalidOperation:
                continue
        # Existing partial rich-text scripts already work and retain their runs.
        if cell.get('t')=='inlineStr' and cell.find('s:is/s:r',NS) is not None:
            continue
        if cell.get('t')=='s' and int(cell.find('s:v',NS).text) in rich_strings:
            continue
        for child in list(cell):
            if child.tag in {f'{{{NS["s"]}}}v',f'{{{NS["s"]}}}is'}:
                cell.remove(child)
        cell.set('t','inlineStr')
        run = ET.SubElement(ET.SubElement(cell,f'{{{NS["s"]}}}is'),f'{{{NS["s"]}}}r')
        properties = ET.SubElement(run,f'{{{NS["s"]}}}rPr')
        for item in font:
            copied = deepcopy(item)
            if copied.tag == f'{{{NS["s"]}}}name':
                copied.tag = f'{{{NS["s"]}}}rFont'
            properties.append(copied)
        value = ET.SubElement(run,f'{{{NS["s"]}}}t')
        value.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        value.text = text
        adjusted.append(cell.get('r'))
    return adjusted
