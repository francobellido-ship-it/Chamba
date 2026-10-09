"""Convert Excel certificates and share the existing PDF publication pipeline."""

from app.excel import convert_excel
from app.pdf import PreparedPDF, prepare_pdf, signature_png


def prepare_document(data: bytes, source_type: str, requested_code: str,
                     verification_url: str, signature: bytes | None = None) -> PreparedPDF:
    if source_type == "xlsx":
        # Validate the optional replacement before starting the Office conversion.
        signature = signature_png(signature)
        converted = convert_excel(data, replace_signature=bool(signature))
        result = prepare_pdf(converted.pdf, requested_code, verification_url, signature, template=True)
        result.signature_included |= converted.embedded_signature
        result.warnings = converted.warnings + result.warnings
        return result
    return prepare_pdf(data, requested_code, verification_url, signature)
