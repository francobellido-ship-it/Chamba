"""Comprueba el portal público con un PDF ficticio."""

import argparse
import hashlib
import json
import os
import urllib.error
import urllib.request
from uuid import uuid4

import pymupdf


def main():
    parser = argparse.ArgumentParser(description="Validar el portal con un certificado ficticio sin validez.")
    parser.add_argument("--url", default=f"http://127.0.0.1:{os.environ.get('PORT', '8000')}")
    args = parser.parse_args()
    base = args.url.rstrip("/")
    with urllib.request.urlopen(base + "/healthz", timeout=10) as response:
        assert json.load(response)["status"] == "ok", "El servicio no está disponible."
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_text((50, 70), "PRUEBA DEL PORTAL - SIN VALIDEZ", fontsize=18)
        page.insert_text((50, 110), "QCP-2026-0001", fontsize=12)
        data = doc.tobytes()
    boundary = uuid4().hex
    body = (f'--{boundary}\r\nContent-Disposition: form-data; name="codigoCorrecto"\r\n\r\n'
            f'QCP-2026-0099\r\n--{boundary}\r\nContent-Disposition: form-data; name="pdf"; '
            'filename="PRUEBA-SIN-VALIDEZ.pdf"\r\nContent-Type: application/pdf\r\n\r\n').encode()
    body += data + f"\r\n--{boundary}--\r\n".encode()
    headers = {"Content-Type": f"multipart/form-data; boundary={boundary}"}
    request = urllib.request.Request(base + "/api/certificates", data=body, headers=headers)
    with urllib.request.urlopen(request, timeout=90) as response:
        assert response.status == 201
        info = json.load(response)
    assert info["finalCode"] == "QCP-2026-0099", "El código no se corrigió."
    with urllib.request.urlopen(base + info["downloadUrl"], timeout=30) as response:
        final_pdf = response.read()
    assert hashlib.sha256(final_pdf).hexdigest() == info["sha256"], "El archivo descargado no coincide."
    with pymupdf.open(stream=final_pdf, filetype="pdf") as doc:
        assert "QCP-2026-0099" in doc[0].get_text()
        assert any(link.get("uri") == info["verificationUrl"] for page in doc for link in page.get_links())
        assert not (doc.permissions & pymupdf.PDF_PERM_MODIFY)
    with urllib.request.urlopen(base + f"/certificados/{info['id']}", timeout=10) as response:
        assert response.status == 200
    print("Portal verificado: carga, corrección, QR enlazado, protección, consulta y descarga.")
    print("Se guardó un certificado ficticio SIN VALIDEZ para esta comprobación.")


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as error:
        raise SystemExit(f"El portal respondió HTTP {error.code}. Revisa la configuración y los registros.")
