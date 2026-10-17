from io import BytesIO

from PIL import Image
import pymupdf
import pytest
import zxingcpp

from app.pdf import prepare_pdf,_template_guides,CertificateError
from app.pdf_qr import instruction_block,verify_template_qr


@pytest.mark.parametrize('gap',[10,16])
def test_visible_qr_is_centered_with_the_complete_instruction_block(gap):
    with pymupdf.open() as source:
        page=source.new_page()
        for i,text in enumerate(['Escanee este QR','Para descargar','el certificado original','con firma electrónica','en pdf']):
            page.insert_text((220,300+i*gap),text,fontsize=8)
        result=prepare_pdf(source.tobytes(),'','https://drive.google.com/file/d/ALINEACION_FICTICIA/view',template=True)
    with pymupdf.open(stream=result.data,filetype='pdf') as doc:
        page=doc[0];block=instruction_block(page,_template_guides(page)[0])
        barcode=zxingcpp.read_barcode(Image.open(BytesIO(page.get_pixmap(matrix=pymupdf.Matrix(4,4)).tobytes('png'))))
        assert barcode.text=='https://drive.google.com/file/d/ALINEACION_FICTICIA/view'
        visible_center=(barcode.position.top_left.y+barcode.position.bottom_left.y)/8
        assert abs(visible_center-(block.y0+block.y1)/2)<1
        assert barcode.position.top_right.x/4<block.x0


def test_unreadable_qr_fails_final_check():
    with pymupdf.open() as doc:
        doc.new_page()
        with pytest.raises(CertificateError):
            verify_template_qr(doc,pymupdf.Rect(10,10,76,76),'https://example.com',b'')
