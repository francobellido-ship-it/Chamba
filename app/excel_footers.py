"""Keep document-control footers outside the scaled certificate body."""

from dataclasses import dataclass
import re

import pymupdf

from app.pdf import CertificateError

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@dataclass(frozen=True)
class FormatFooter:
    code: str
    version: str


def _plain(text):
    return re.sub(r'&"[^\"]*"|&K[0-9A-Fa-f]{6}|&[0-9]+|&[BIOUSENXY]', '', text)


def take_format_footers(sheet):
    """Remove only recognized control fields; leave other native footer text."""
    settings = {}
    header = sheet.find("s:headerFooter", NS)
    if header is None:
        return settings
    for element in header:
        variant = element.tag.rsplit("}", 1)[-1]
        if not variant.endswith("Footer"):
            continue
        text = element.text or ""
        code = re.search(r'C[oó]digo\s*:\s*(F-GCA-[0-9.]+)', _plain(text), re.I)
        version = re.search(r'(?:Versi[oó]n|Revisi[oó]n)\s*:\s*([0-9A-Za-z.-]+)', _plain(text), re.I)
        if code and version:
            settings[variant.removesuffix("Footer")] = FormatFooter(code[1].upper(), version[1])
            text = re.sub(r'C[oó]digo\s*:\s*F-GCA-[0-9.]+', '', text, flags=re.I)
            text = re.sub(r'(?:Versi[oó]n|Revisi[oó]n)\s*:\s*[0-9A-Za-z.-]+', '', text, flags=re.I)
        text = re.sub(r'P[aá]g(?:ina|\.)?\s*&P(?:\s*de\s*&N)?', '', text, flags=re.I)
        element.text = text
    return settings


def footer_for_page(settings, index, headers):
    first, even = headers[2:]
    variant = 'first' if first and index == 0 else 'even' if even and (index+1) % 2 == 0 else 'odd'
    return settings.get(variant) or settings.get('odd')


def apply_format_footers(doc, footers):
    font = pymupdf.Font('helv')
    for index, (page, footer) in enumerate(zip(doc, footers, strict=True)):
        lines = [f'Código: {footer.code}', f'Versión: {footer.version}'] if footer else []
        lines.append(f'Página {index+1} de {len(doc)}')
        # Keep one compact, evenly spaced block in the reserved lower band.
        # Anchor the page number so missing format fields do not move it.
        for offset, text in enumerate(lines):
            width = font.text_length(text, fontsize=5)
            baseline = page.rect.height-10-(len(lines)-1-offset)*8
            page.insert_text((page.rect.width-width-7, baseline), text, fontsize=5)
    verify_format_footers(doc, footers)


def verify_format_footers(doc, footers):
    for index, (page, footer) in enumerate(zip(doc, footers, strict=True)):
        texts = [f'Código: {footer.code}', f'Versión: {footer.version}'] if footer else []
        for text in texts + [f'Página {index+1} de {len(doc)}']:
            areas = page.search_for(text, clip=pymupdf.Rect(0,page.rect.height-44,page.rect.width,page.rect.height))
            if len(areas) != 1 or not page.rect.contains(areas[0]) or areas[0].y0 < page.rect.height-44:
                raise CertificateError(f'No se pudo verificar el código, versión y página del pie en la página {index+1}.')


def entered_footer(code, version):
    code, version = code.strip().upper(), version.strip()
    if not code and not version:
        return None
    if not re.fullmatch(r'[A-Z0-9][A-Z0-9.-]{2,39}', code) or not re.fullmatch(r'[0-9A-Za-z][0-9A-Za-z._-]{0,15}', version):
        raise CertificateError('Completa el código y la versión del formato, o deja ambos vacíos para leerlos del Excel.')
    return FormatFooter(code, version)


def general_data_footer(parts):
    from app.excel import xml, cell_value, R, related
    book = xml(parts['xl/workbook.xml'])
    strings = [''.join(item.itertext()) for item in xml(parts['xl/sharedStrings.xml'])] if 'xl/sharedStrings.xml' in parts else []
    targets = {rel.get('Id'):related('xl/workbook.xml',rel.get('Target','')) for rel in xml(parts['xl/_rels/workbook.xml.rels']) if rel.get('TargetMode') != 'External'}
    styles = xml(parts['xl/styles.xml'])
    formats = {int(fmt.get('numFmtId')):fmt.get('formatCode','') for fmt in styles.findall('s:numFmts/s:numFmt',NS)}
    xfs = styles.findall('s:cellXfs/s:xf',NS)
    def saved_text(cell):
        value = cell_value(cell,strings).strip()
        style = int(cell.get('s','0'))
        code = formats.get(int(xfs[style].get('numFmtId','0')),'') if style < len(xfs) else ''
        return value.zfill(len(code)) if value.isdigit() and re.fullmatch('0{2,4}',code) else value
    found = set()
    for entry in book.findall('s:sheets/s:sheet',NS):
        if not re.search(r'datos|generales|informaci[oó]n|registro|inicio',entry.get('name',''),re.I):
            continue
        sheet = xml(parts[targets[entry.get(f'{{{R}}}id')]])
        cells = {cell.get('r'):saved_text(cell) for cell in sheet.findall('s:sheetData/s:row/s:c',NS)}
        for ref, value in cells.items():
            if not re.fullmatch(r'F-GCA-10\.[0-9.]+',value,re.I):
                continue  # Internal register/calculation form IDs are different documents.
            column, row = re.fullmatch(r'([A-Z]+)([0-9]+)',ref).groups()
            for offset in range(-2,4):
                candidate = cells.get(f'{column}{int(row)+offset}','')
                labelled = re.fullmatch(r'(?:Versi[oó]n|Ver\.?|Revisi[oó]n)\s*:?\s*([0-9A-Za-z.-]+)',candidate,re.I)
                if labelled:
                    found.add(FormatFooter(value.upper(),labelled[1]))
                elif offset and re.fullmatch(r'[0-9]{1,3}',candidate):
                    labels = [text for key,text in cells.items() if re.fullmatch(r'[A-Z]+'+str(int(row)+offset),key)
                              and re.fullmatch(r'(?:Versi[oó]n|Revisi[oó]n)(?:\s+del\s+formato)?\s*:?\s*',text,re.I)]
                    if labels:
                        found.add(FormatFooter(value.upper(),candidate))
    return next(iter(found)) if len(found)==1 else None


def inspect_format_footer(data):
    from app.excel import _read_parts,xml
    from app.excel_sheets import certificate_layouts
    parts = _read_parts(data)
    layout = certificate_layouts(data)[0]
    own = take_format_footers(xml(parts[layout.sheet_path]))
    footer = own.get('odd') or next(iter(own.values()),None) or general_data_footer(parts)
    return footer
