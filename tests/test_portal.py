from io import BytesIO
import hashlib
from uuid import uuid4

from fastapi.testclient import TestClient
from PIL import Image
import pymupdf
import pytest
import zxingcpp

from app.main import create_app
from app.pdf import CertificateError, prepare_pdf


def pdf_bytes(codes=("QCP-2026-0001",), *, full_page=False):
    with pymupdf.open() as doc:
        for code in codes:
            page = doc.new_page()
            page.insert_text((50, 60), "Certificado de ensayo", fontsize=18)
            page.insert_text((50, 100), code, fontsize=12)
            page.insert_text((50, 140), "Resultado: conforme. Contenido que debe conservarse.", fontsize=11)
            if full_page:
                page.draw_rect(page.rect, color=(0.5, 0.5, 0.5))
        return doc.tobytes()


@pytest.fixture
def client(tmp_path):
    with TestClient(create_app(tmp_path / "certificates", "http://127.0.0.1:8000")) as client:
        yield client


def upload(client, data=None, **fields):
    return client.post("/api/certificates", files={"pdf": ("certificado.pdf", data or pdf_bytes(), "application/pdf")}, data=fields)


def test_home_config_and_health(client):
    assert client.get("/healthz").json() == {"status": "ok"}
    response = client.get("/")
    assert response.status_code == 200
    assert "Preparar certificados" in response.text
    assert "frame-ancestors 'none'" in response.headers["content-security-policy"]
    config = client.get("/api/config").json()
    assert config["localMode"] is True
    assert config["requiresAccessKey"] is False


def test_upload_replace_and_download_end_to_end(client):
    response = upload(client, pdf_bytes(("QCP-2026-0001", "QCP-2026-0001")), codigoCorrecto="qcp-2026-0099")
    assert response.status_code == 201, response.text
    info = response.json()
    assert info["finalCode"] == "QCP-2026-0099"
    assert info["replacements"] == 2
    assert len(info["stampPages"]) == 2
    assert client.get(f"/certificados/{info['id']}").status_code == 200
    assert client.get(f"/api/certificates/{info['id']}").json() == info
    download = client.get(info["downloadUrl"])
    assert download.status_code == 200
    assert download.headers["content-type"] == "application/pdf"
    assert "certificado.pdf" in download.headers["content-disposition"]
    assert hashlib.sha256(download.content).hexdigest() == info["sha256"]
    with pymupdf.open(stream=download.content, filetype="pdf") as doc:
        assert not doc.needs_pass
        assert doc.metadata["encryption"]
        assert not (doc.permissions & pymupdf.PDF_PERM_MODIFY)
        assert doc.permissions & pymupdf.PDF_PERM_PRINT
        for page in doc:
            text = page.get_text()
            assert "QCP-2026-0099" in text
            assert "QCP-2026-0001" not in text
            assert "Resultado: conforme" in text
            assert any(link.get("uri") == info["verificationUrl"] for link in page.get_links())
        # Decode the rendered QR, including its transparent margin against the page.
        decoded = []
        for image in doc[0].get_image_info():
            pixmap = doc[0].get_pixmap(matrix=pymupdf.Matrix(3, 3), clip=pymupdf.Rect(image["bbox"]))
            barcode = zxingcpp.read_barcode(Image.open(BytesIO(pixmap.tobytes("png"))))
            if barcode:
                decoded.append(barcode.text)
        assert info["verificationUrl"] in decoded


def test_optional_signature(client):
    image = Image.new("RGBA", (240, 80), (0, 50, 80, 200))
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    response = client.post("/api/certificates", files={
        "pdf": ("firma.pdf", pdf_bytes(), "application/pdf"),
        "firma": ("firma.png", buffer.getvalue(), "image/png")})
    assert response.status_code == 201, response.text
    assert response.json()["signatureIncluded"] is True
    with pymupdf.open(stream=client.get(response.json()["downloadUrl"]).content, filetype="pdf") as doc:
        assert len(doc[0].get_images()) == 2
        assert "Firma autorizada" in doc[0].get_text()


def test_multiple_codes_are_not_changed(client):
    response = upload(client, pdf_bytes(("QCP-2026-0001", "QCP-2026-0002")), codigoCorrecto="QCP-2026-0099")
    assert response.status_code == 201
    info = response.json()
    assert info["finalCode"] is None
    assert info["replacements"] == 0
    assert len(info["detectedCodes"]) == 2
    assert any("varios códigos" in warning for warning in info["warnings"])
    with pymupdf.open(stream=client.get(info["downloadUrl"]).content, filetype="pdf") as doc:
        assert "QCP-2026-0001" in doc[0].get_text()
        assert "QCP-2026-0002" in doc[1].get_text()


def test_no_code_does_not_claim_a_requested_replacement(client):
    response = upload(client, pdf_bytes(("Sin código",)), codigoCorrecto="QCP-2026-0099")
    assert response.status_code == 201
    assert response.json()["finalCode"] is None
    assert response.json()["warnings"]


def test_no_space_adds_page_without_covering_existing_content(client):
    response = upload(client, pdf_bytes(full_page=True))
    assert response.status_code == 201
    info = response.json()
    assert info["stampPages"] == [2]
    assert any("página adicional" in warning for warning in info["warnings"])
    with pymupdf.open(stream=client.get(info["downloadUrl"]).content, filetype="pdf") as doc:
        assert len(doc) == 2
        assert "Resultado: conforme" in doc[0].get_text()
        assert not doc[0].get_images()
        assert len(doc[1].get_images()) == 1


@pytest.mark.parametrize("data,code,expected", [
    (b"not a pdf", "", "no es un PDF"),
    (b"%PDF-1.7 broken", "", "dañado"),
    (pdf_bytes(), "QCP-123", "formato"),
    (pdf_bytes(), "QCP-٢٠٢٦-٠٠٠١", "formato"),
])
def test_invalid_input_returns_actionable_error(client, data, code, expected):
    response = upload(client, data, codigoCorrecto=code)
    assert response.status_code == 422
    assert expected in response.json()["detail"]


def test_invalid_signature_does_not_store_certificate(client, tmp_path):
    response = client.post("/api/certificates", files={
        "pdf": ("test.pdf", pdf_bytes(), "application/pdf"),
        "firma": ("firma.png", b"not a picture", "image/png")})
    assert response.status_code == 422
    assert not list((tmp_path / "certificates").iterdir())


def test_password_protected_pdf_is_rejected(client):
    with pymupdf.open(stream=pdf_bytes(), filetype="pdf") as doc:
        encrypted = doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="owner", user_pw="locked")
    response = upload(client, encrypted)
    assert response.status_code == 422
    assert "contraseña" in response.json()["detail"]


def test_signature_field_is_not_silently_invalidated():
    with pymupdf.open(stream=pdf_bytes(), filetype="pdf") as doc:
        widget = pymupdf.Widget()
        widget.field_type = pymupdf.PDF_WIDGET_TYPE_SIGNATURE
        widget.field_name = "Firma digital"
        widget.rect = pymupdf.Rect(50, 250, 250, 300)
        doc[0].add_widget(widget)
        data = doc.tobytes()
    with pytest.raises(CertificateError, match="firma digital"):
        prepare_pdf(data, "", "https://example.com/certificados/test")


def test_public_deployment_requires_https(tmp_path):
    with pytest.raises(ValueError, match="HTTPS"):
        create_app(tmp_path, "http://example.com")


@pytest.mark.parametrize("key", ["", "short", "test-key-at-least-16-chars"])
def test_public_uploads_ignore_legacy_access_key(tmp_path, monkeypatch, key):
    monkeypatch.setenv("PORTAL_ACCESS_KEY", key)
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        assert client.get("/").status_code == 200
        assert client.get("/healthz").json() == {"status": "ok"}
        assert client.get("/api/config").json()["uploadsEnabled"] is True
        assert client.get("/api/config").json()["requiresAccessKey"] is False
        response = upload(client, codigoCorrecto="QCP-2026-0099")
        assert response.status_code == 201
        assert response.json()["finalCode"] == "QCP-2026-0099"
        assert client.get(response.json()["downloadUrl"]).status_code == 200


def test_render_url_without_access_key_serves_public_page(tmp_path, monkeypatch):
    monkeypatch.setenv("RENDER_EXTERNAL_URL", "https://qcp-example.onrender.com")
    monkeypatch.delenv("PUBLIC_BASE_URL", raising=False)
    monkeypatch.delenv("PORTAL_ACCESS_KEY", raising=False)
    with TestClient(create_app(storage_dir=tmp_path)) as client:
        assert client.get("/").status_code == 200
        assert client.get("/api/config").json()["localMode"] is False
        assert client.get("/api/config").json()["uploadsEnabled"] is True
        assert upload(client).status_code == 201


def test_existing_certificates_remain_public_after_restart(tmp_path):
    with TestClient(create_app(tmp_path, "https://example.com")) as configured:
        response = upload(configured)
        assert response.status_code == 201
        info = response.json()
    with TestClient(create_app(tmp_path, "https://example.com")) as unconfigured:
        assert unconfigured.get(f"/certificados/{info['id']}").status_code == 200
        assert unconfigured.get(f"/api/certificates/{info['id']}").json() == info
        assert hashlib.sha256(unconfigured.get(info["downloadUrl"]).content).hexdigest() == info["sha256"]


def test_public_uploads_do_not_require_an_auth_header(tmp_path):
    with TestClient(create_app(tmp_path, "https://certificates.example.com")) as client:
        response = client.post("/api/certificates", headers={"X-Portal-Key": "irrelevant-old-header"},
                               files={"pdf": ("test.pdf", pdf_bytes(), "application/pdf")})
        assert response.status_code == 201
        assert response.json()["verificationUrl"].startswith("https://certificates.example.com/")
        assert client.get(response.json()["downloadUrl"]).status_code == 200
        assert client.get("/api/config").json()["requiresAccessKey"] is False
        assert client.get("/api/config").json()["uploadsEnabled"] is True


def test_missing_and_malformed_certificate_id(client):
    assert client.get(f"/api/certificates/{uuid4()}").status_code == 404
    assert client.get("/api/certificates/not-an-id/download").status_code == 422


def test_oversized_request_is_rejected_before_processing(client):
    response = client.post("/api/certificates", content=b"test", headers={"Content-Length": str(19 * 1024 * 1024)})
    assert response.status_code == 413


def test_pdf_file_limit_is_enforced(client):
    response = upload(client, b"%PDF-1.7\n" + b"0" * (16 * 1024 * 1024))
    assert response.status_code == 413
    assert "15 MB" in response.json()["detail"]


def test_registered_certificate_survives_application_restart(tmp_path):
    with TestClient(create_app(tmp_path, "http://127.0.0.1:8000")) as first:
        response = upload(first)
        assert response.status_code == 201
        info = response.json()
    with TestClient(create_app(tmp_path, "http://127.0.0.1:8000")) as restarted:
        download = restarted.get(info["downloadUrl"])
        assert download.status_code == 200
        assert hashlib.sha256(download.content).hexdigest() == info["sha256"]


def test_storage_quota_does_not_silently_remove_files(tmp_path, monkeypatch):
    monkeypatch.setenv("MAX_STORAGE_MB", "0")
    with TestClient(create_app(tmp_path, "http://127.0.0.1:8000")) as client:
        response = upload(client)
        assert response.status_code == 507
        assert not list(tmp_path.iterdir())
