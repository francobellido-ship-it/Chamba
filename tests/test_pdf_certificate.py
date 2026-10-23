"""PDF-only flow: background and guided QR, with original content untouched."""

from io import BytesIO
import hashlib

from fastapi.testclient import TestClient
from PIL import Image
import pymupdf
import pytest
import qrcode
import zxingcpp

from app.backgrounds import load_letterhead,verify_letterhead
from app.documents import prepare_document
from app.pdf import CertificateError
from app.main import create_app
from test_backgrounds import synthetic_backgrounds
from test_drive import FakeDrive,ROOT


def certificate_pdf(*,existing_qr=False,background=None,blocked=False,guide=True,shared_qr=False):
    image=BytesIO();Image.new('RGB',(101,51),'navy').save(image,format='PNG')
    old_qr=BytesIO();qrcode.make('https://example.com/old-certificate').save(old_qr,format='PNG')
    with pymupdf.open()as doc:
        for index,(width,height)in enumerate([(595.276,841.89),(612,792)]):
            page=doc.new_page(width=width,height=height)
            if background:
                page.insert_image(page.rect,stream=background.data,overlay=False,keep_proportion=False)
            page.insert_text((80,150),'FICTICIO - SIN VALIDEZ')
            page.insert_text((80,180),'QCP-2026-0001')
            page.insert_text((80,210),f'RESULTADO GUARDADO: 17,50. Pagina {index+1} de 2')
            page.draw_rect(pymupdf.Rect(70,225,500,260))
            page.insert_text((80,245),'TABLA ORIGINAL: 0,00000000001')
            page.insert_image(pymupdf.Rect(350,625,451,676),stream=image.getvalue())
            page.insert_text((430,770),'F-GCA-10.88 / Version 07',fontsize=6)
            page.insert_link({'kind':pymupdf.LINK_URI,'from':pymupdf.Rect(80,300,200,320),'uri':'https://example.com/keep-link'})
            if index==0 and guide:
                page.insert_text((220,650),'Escanee este QR',fontsize=9)
                page.insert_text((220,661),'Para descargar el certificado',fontsize=9)
                page.insert_text((220,672),'con firma electronica',fontsize=9)
            if index==0 and blocked:
                page.draw_line((150,640),(210,670))
            if existing_qr and (index==0 or shared_qr):
                page.insert_image(pymupdf.Rect(148,630,214,696),stream=old_qr.getvalue())
                page.insert_link({'kind':pymupdf.LINK_URI,'from':pymupdf.Rect(148,630,214,696),'uri':'https://example.com/old-certificate'})
        return doc.tobytes()


@pytest.mark.parametrize('kind',['acreditado','no_acreditado'])
@pytest.mark.parametrize('existing_qr',[False,True])
def test_pdf_keeps_words_positions_tables_images_links_and_pages(synthetic_backgrounds,kind,existing_qr):
    selected=load_letterhead(kind)
    source=certificate_pdf(existing_qr=existing_qr,background=selected)
    url='https://example.com/new-certificate'
    result=prepare_document(source,'pdf','QCP-2026-9999',url,b'invalid-signature',certificate_type=kind)
    assert result.final_code=='QCP-2026-0001' and result.replacements==0
    assert result.stamp_pages==[1] and not result.signature_included
    assert result.background_pages==[1,2] and result.certificate_type==kind
    with pymupdf.open(stream=source,filetype='pdf')as old,pymupdf.open(stream=result.data,filetype='pdf')as new:
        assert len(old)==len(new)==2
        assert verify_letterhead(new,selected,preserve_page_size=True)==[1,2]
        for before,after in zip(old,new,strict=True):
            assert before.rect==after.rect and before.rotation==after.rotation
            assert before.get_text('words')==after.get_text('words')
            assert before.get_drawings()==after.get_drawings()
            signature=lambda p:[(i['width'],i['height'],i['bbox'],i['digest'])
                                for i in p.get_image_info(hashes=True)if i['width']==101]
            assert signature(before)==signature(after)
            assert any(link['uri']=='https://example.com/keep-link'for link in after.get_links())
        assert len([l for p in new for l in p.get_links()if l.get('uri')==url])==1
        assert not any(l.get('uri')=='https://example.com/old-certificate'for l in new[0].get_links())
        # A page with an unchanged background and no new QR remains pixel-identical.
        assert old[1].get_pixmap().samples==new[1].get_pixmap().samples
        assert not new.metadata['encryption'] and new.permissions&pymupdf.PDF_PERM_MODIFY


@pytest.mark.parametrize('kind',['acreditado','no_acreditado'])
def test_pdf_can_change_official_background_without_doubling_it(synthetic_backgrounds,kind):
    other=load_letterhead('no_acreditado'if kind=='acreditado'else'acreditado')
    source=certificate_pdf(background=other)
    result=prepare_document(source,'pdf','','https://example.com/new',certificate_type=kind)
    with pymupdf.open(stream=source,filetype='pdf')as before,pymupdf.open(stream=result.data,filetype='pdf')as after:
        assert verify_letterhead(after,load_letterhead(kind),preserve_page_size=True)==[1,2]
        for old,page in zip(before,after,strict=True):
            assert page.get_text('words')==old.get_text('words')
            assert not any(i['digest']==other.digest for i in page.get_image_info(hashes=True))


@pytest.mark.parametrize('options,match',[
    ({'guide':False},'única indicación'),
    ({'blocked':True},'línea o un cuadro'),
    ({'existing_qr':True,'shared_qr':True},'comparte entre varias'),
])
def test_ambiguous_or_occupied_qr_slot_fails_without_reflow(synthetic_backgrounds,options,match):
    with pytest.raises(CertificateError,match=match):
        prepare_document(certificate_pdf(**options),'pdf','','https://example.com/new',certificate_type='no_acreditado')


def test_pdf_api_uses_only_background_and_qr_even_when_other_fields_are_sent(tmp_path,synthetic_backgrounds):
    with TestClient(create_app(tmp_path/'store','https://example.com'))as client:
        response=client.post('/api/certificates',files={
            'pdf':('certificate.pdf',certificate_pdf(),'application/pdf'),
            'firma':('invalid.png',b'not-an-image','image/png')},
            data={'tipoCertificado':'acreditado','codigoCorrecto':'not-a-certificate-code',
                  'codigoFormato':'F-GCA-10.99','versionFormato':'99'})
        assert response.status_code==201,response.text
        info=response.json()
        assert info['sourceType']=='pdf' and info['backgroundPages']==[1,2]
        assert info['finalCode']=='QCP-2026-0001' and not info['signatureIncluded']
        assert info['certificateType']=='acreditado' and info['replacements']==0
        with pymupdf.open(stream=client.get(info['downloadUrl']).content,filetype='pdf')as doc:
            assert 'F-GCA-10.88 / Version 07'in doc[0].get_text()
            assert 'F-GCA-10.99'not in doc[0].get_text()


def test_invalid_pdf_background_does_not_allocate_drive_publication():
    drive=FakeDrive()
    with TestClient(create_app(certificate_storage=drive.storage()))as client:
        response=client.post('/api/certificates',files={'pdf':('test.pdf',certificate_pdf())},
                             data={'tipoCertificado':'unknown'})
    assert response.status_code==422 and drive.token_requests==0 and list(drive.files)==[ROOT]


@pytest.mark.parametrize('fill,opacity',[( (1,1,1),1),((0,0,0),0)])
def test_empty_cell_paint_does_not_block_qr_or_change_original_content(synthetic_backgrounds,fill,opacity):
    with pymupdf.open(stream=certificate_pdf(),filetype='pdf')as doc:
        doc[0].draw_rect(pymupdf.Rect(140,620,216,710),color=None,fill=fill,fill_opacity=opacity,overlay=False)
        source=doc.tobytes()
    result=prepare_document(source,'pdf','','https://example.com/new',certificate_type='no_acreditado')
    with pymupdf.open(stream=source,filetype='pdf')as old,pymupdf.open(stream=result.data,filetype='pdf')as new:
        for before,after in zip(old,new,strict=True):
            assert before.get_text('words')==after.get_text('words')
            shapes=lambda page:[{k:v for k,v in d.items()if k!='seqno'}for d in page.get_drawings()]
            assert shapes(before)==shapes(after)
        assert len(new)==2 and result.stamp_pages==[1]


@pytest.mark.parametrize('obstacle',['colored_fill','border','text','image'])
def test_white_cell_still_protects_content_and_borders(synthetic_backgrounds,obstacle):
    with pymupdf.open(stream=certificate_pdf(),filetype='pdf')as doc:
        page=doc[0];area=pymupdf.Rect(150,640,210,670)
        page.draw_rect(area,color=None,fill=(1,1,1))
        if obstacle=='colored_fill':
            page.draw_rect(area,color=None,fill=(.8,.8,.8))
        elif obstacle=='border':
            page.draw_rect(area,color=(0,0,0),fill=(1,1,1))
        elif obstacle=='text':
            page.insert_text((155,660),'KEEP',fontsize=8)
        else:
            image=BytesIO();Image.new('RGB',(60,30),'navy').save(image,format='PNG')
            page.insert_image(area,stream=image.getvalue())
        source=doc.tobytes()
    with pytest.raises(CertificateError):
        prepare_document(source,'pdf','','https://example.com/new',certificate_type='no_acreditado')


@pytest.mark.parametrize('backend',['local','drive'])
def test_republishing_same_pdf_keeps_each_qr_and_previous_file(tmp_path,synthetic_backgrounds,backend):
    drive=FakeDrive()
    app=(create_app(certificate_storage=drive.storage())if backend=='drive'
         else create_app(tmp_path/'store','https://example.com'))
    source=certificate_pdf();records=[]
    with TestClient(app)as client:
        for _ in range(2):
            response=client.post('/api/certificates',files={'pdf':('same-name.pdf',source,'application/pdf')},
                                 data={'tipoCertificado':'no_acreditado'})
            assert response.status_code==201,response.text
            info=response.json();data=client.get(info['downloadUrl']).content
            assert hashlib.sha256(data).hexdigest()==info['sha256']
            with pymupdf.open(stream=data,filetype='pdf')as doc:
                area=[l['from']for l in doc[0].get_links()if l.get('uri')==info['verificationUrl']][0]
                pixels=doc[0].get_pixmap(matrix=pymupdf.Matrix(3,3),clip=area)
                decoded=zxingcpp.read_barcode(Image.open(BytesIO(pixels.tobytes('png'))))
                assert decoded.text==info['verificationUrl']
            records.append((info,data))
        first,second=(row[0]for row in records)
        assert first['id']!=second['id'] and first['verificationUrl']!=second['verificationUrl']
        assert first['filename']==second['filename'] and first['finalCode']==second['finalCode']
        for info,data in records:
            assert client.get(info['downloadUrl']).content==data
            assert client.get(f"/api/certificates/{info['id']}").json()['sha256']==info['sha256']
        if backend=='drive':
            assert first['driveFileId']!=second['driveFileId']
            assert all(drive.files[info['driveFileId']]['content']==data for info,data in records)
