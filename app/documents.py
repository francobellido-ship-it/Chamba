"""Convert Excel certificates and share the existing PDF publication pipeline."""

from app.excel import convert_excel
from app.pdf import PreparedPDF, prepare_pdf, signature_png
from app.backgrounds import load_letterhead, verify_letterhead
import pymupdf


def prepare_document(data: bytes, source_type: str, requested_code: str,
                     verification_url: str, signature: bytes | None = None,
                     certificate_type: str = "") -> PreparedPDF:
    if source_type == "xlsx":
        # Validate the optional replacement before starting the Office conversion.
        signature = signature_png(signature)
        letterhead = load_letterhead(certificate_type)
        converted = convert_excel(data, replace_signature=bool(signature), letterhead=letterhead)
        result = prepare_pdf(converted.pdf, requested_code, verification_url, signature, template=True)
        result.signature_included |= converted.embedded_signature
        result.warnings = converted.warnings + result.warnings
        # Validate the final encrypted PDF after QR, signature and code processing.
        with pymupdf.open(stream=result.data, filetype="pdf") as doc:
            result.background_pages = verify_letterhead(doc, letterhead)
        result.certificate_type = certificate_type
        return result
    return prepare_pdf(data, requested_code, verification_url, signature)
