from contextlib import asynccontextmanager
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timezone
import asyncio
import hashlib
import ipaddress
import logging
import multiprocessing
import os
from pathlib import Path
import re
from urllib.parse import quote, urlsplit
from uuid import UUID, uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException

from app.pdf import CertificateError
from app.documents import prepare_document
from app.excel import inspect_excel
from app.storage import DriveStorage, LocalStorage, StorageError, storage_from_environment
from app.backgrounds import background_choices, require_background_type

STATIC = Path(__file__).parent / "static"
MAX_PDF_BYTES = 15 * 1024 * 1024
MAX_SIGNATURE_BYTES = 2 * 1024 * 1024
MAX_REQUEST_BYTES = 18 * 1024 * 1024
logger = logging.getLogger(__name__)


def local_url(url: str) -> bool:
    hostname = urlsplit(url).hostname
    if hostname == "localhost":
        return True
    try:
        return ipaddress.ip_address(hostname or "").is_loopback
    except ValueError:
        return False


class RequestGuard:
    """Bound streamed upload request bodies before processing certificates."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)
        is_upload = scope["method"] == "POST" and scope["path"] in {"/api/certificates", "/api/excel/inspect"}
        if not is_upload:
            return await self.app(scope, receive, send)
        headers = dict(scope["headers"])
        try:
            length = int(headers.get(b"content-length", b"0"))
        except ValueError:
            return await JSONResponse({"detail": "Tamaño de solicitud inválido."}, status_code=400)(scope, receive, send)
        if length > MAX_REQUEST_BYTES:
            return await JSONResponse({"detail": "La carga supera el límite permitido."}, status_code=413)(scope, receive, send)
        received = 0

        async def bounded_receive():
            nonlocal received
            message = await receive()
            received += len(message.get("body", b""))
            if received > MAX_REQUEST_BYTES:
                raise StarletteHTTPException(413, "La carga supera el límite permitido.")
            return message

        await self.app(scope, bounded_receive, send)


async def read_upload(file: UploadFile, limit: int) -> bytes:
    chunks = []
    count = 0
    try:
        while chunk := await file.read(64 * 1024):
            count += len(chunk)
            if count > limit:
                raise HTTPException(413, f"El archivo supera el límite de {limit // (1024 * 1024)} MB.")
            chunks.append(chunk)
    finally:
        await file.close()
    return b"".join(chunks)


def filename_safe(name: str | None, source_type: str = "pdf") -> str:
    name = (name or "CERTIFICADO.pdf").replace("\\", "/").split("/")[-1]
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip()
    if not name.lower().endswith(f".{source_type}"):
        raise HTTPException(422, f"El archivo debe tener extensión .{source_type}.")
    return name[:-(len(source_type) + 1)][:176] + ".pdf"


def create_app(storage_dir: Path | None = None, public_url: str | None = None,
               certificate_storage: LocalStorage | DriveStorage | None = None) -> FastAPI:
    storage = storage_dir or Path(os.environ.get("CERTIFICATE_STORAGE_DIR", "data/certificates"))
    origin = (public_url or os.environ.get("PUBLIC_BASE_URL") or os.environ.get("RENDER_EXTERNAL_URL")
              or "http://127.0.0.1:8000").rstrip("/")
    parsed = urlsplit(origin)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.path or parsed.query or parsed.fragment or parsed.username or parsed.password:
        raise ValueError("PUBLIC_BASE_URL debe ser una URL http(s) sin ruta, credenciales, consulta ni fragmento.")
    is_local = local_url(origin)
    if not is_local and parsed.scheme != "https":
        raise ValueError("PUBLIC_BASE_URL debe usar HTTPS para publicar el portal.")
    quota = int(os.environ.get("MAX_STORAGE_MB", "1024")) * 1024 * 1024
    store = certificate_storage or (LocalStorage(storage, quota) if storage_dir else storage_from_environment(storage, quota))

    @asynccontextmanager
    async def lifespan(app):
        await asyncio.to_thread(store.start)
        app.state.executor = ProcessPoolExecutor(max_workers=1, mp_context=multiprocessing.get_context("spawn"))
        app.state.processing = asyncio.Semaphore(1)
        yield
        app.state.executor.shutdown(wait=True, cancel_futures=True)
        await asyncio.to_thread(store.close)

    app = FastAPI(title="Portal de certificados QCP", lifespan=lifespan,
                  docs_url=None, redoc_url=None, openapi_url=None)
    app.add_middleware(RequestGuard)

    @app.exception_handler(StorageError)
    async def storage_error(request, error):
        return JSONResponse({"detail": str(error)}, status_code=error.status)

    @app.middleware("http")
    async def security_headers(request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; img-src 'self' data: blob:; script-src 'self'; "
            "style-src 'self'; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
        )
        if request.url.path.startswith("/api/"):
            response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/healthz")
    async def health():
        return {"status": "ok"}

    @app.get("/api/config")
    async def configuration():
        return {"requiresAccessKey": False, "uploadsEnabled": store.ready,
                "localMode": is_local and store.backend == "local", "storageBackend": store.backend,
                "maxPdfMB": 15, "maxExcelMB": 15, "maxSignatureMB": 2, "maxPages": 100,
                "inputFormats": ["xlsx", "pdf"], "certificateBackgrounds": background_choices()}

    @app.get("/")
    async def home():
        return FileResponse(STATIC / "index.html")

    @app.get("/privacidad")
    async def privacy():
        return FileResponse(STATIC / "privacy.html")

    @app.get("/condiciones")
    async def terms():
        return FileResponse(STATIC / "terms.html")

    @app.get("/certificados/{certificate_id}")
    async def certificate_page(certificate_id: UUID):
        await asyncio.to_thread(store.metadata, str(certificate_id))
        return FileResponse(STATIC / "verify.html", headers={"Cache-Control": "no-store"})

    @app.get("/api/certificates/{certificate_id}")
    async def details(certificate_id: UUID):
        return await asyncio.to_thread(store.metadata, str(certificate_id))

    @app.get("/api/certificates/{certificate_id}/download")
    async def download(certificate_id: UUID):
        info = await asyncio.to_thread(store.metadata, str(certificate_id))
        if isinstance(store, LocalStorage):
            return FileResponse(store.local_pdf(str(certificate_id)), media_type="application/pdf",
                                filename=info["filename"], headers={"Cache-Control": "no-store"})
        pdf = await asyncio.to_thread(store.download, info)
        return Response(pdf, media_type="application/pdf", headers={"Cache-Control": "no-store",
            "Content-Disposition": f"attachment; filename*=utf-8''{quote(info['filename'])}"})

    @app.post("/api/certificates", status_code=201)
    async def create_certificate(excel: UploadFile | None = File(None), pdf: UploadFile | None = File(None), codigoCorrecto: str = Form(""),
                                 firma: UploadFile | None = File(None), tipoCertificado: str = Form("")):
        if not store.ready:
            raise StorageError("Falta completar la conexión con Google Drive. Contacta al administrador.")
        if (excel is None) == (pdf is None):
            raise HTTPException(422, "Sube un Excel .xlsx o un PDF por solicitud.")
        source_type = "xlsx" if excel else "pdf"
        document = excel or pdf
        filename = filename_safe(document.filename, source_type)
        if len(codigoCorrecto) > 40:
            raise HTTPException(422, "El código debe tener el formato QCP-0000-0000.")
        data = await read_upload(document, MAX_PDF_BYTES)
        if source_type == "xlsx":
            try:
                layout = await asyncio.to_thread(inspect_excel, data)
                if not layout.preserve_format:
                    require_background_type(tipoCertificado)
            except CertificateError as exc:
                raise HTTPException(422, str(exc)) from exc
        signature = await read_upload(firma, MAX_SIGNATURE_BYTES) if firma else None
        certificate_id = str(uuid4())
        try:
            await asyncio.wait_for(app.state.processing.acquire(), timeout=0.1)
        except TimeoutError:
            raise HTTPException(429, "El portal está procesando otros archivos. Inténtalo nuevamente en unos segundos.")
        try:
            publication = await asyncio.to_thread(store.allocate, certificate_id, origin)
            verification_url = publication.verification_url
            result = await asyncio.get_running_loop().run_in_executor(
                app.state.executor, prepare_document, data, source_type, codigoCorrecto, verification_url, signature, tipoCertificado)
            info = {"id": certificate_id, "filename": filename, "sourceType": source_type,
                    "createdAt": datetime.now(timezone.utc).isoformat(),
                    "detectedCodes": result.detected_codes, "finalCode": result.final_code,
                    "replacements": result.replacements, "warnings": result.warnings,
                    "stampPages": result.stamp_pages, "signatureIncluded": result.signature_included,
                    "sha256": hashlib.sha256(result.data).hexdigest(),
                    "verificationUrl": verification_url,
                    "downloadUrl": f"/api/certificates/{certificate_id}/download"}
            if source_type == "xlsx":
                info.update(certificateType=result.certificate_type, backgroundVerified=bool(result.background_pages),
                            backgroundPages=result.background_pages, formatPreserved=result.format_preserved,
                            sheetName=result.sheet_name)
            await asyncio.to_thread(store.save, certificate_id, result.data, info, publication)
        except StorageError:
            raise
        except CertificateError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            logger.exception("Error al procesar un certificado")
            raise HTTPException(500, "No se pudo procesar el certificado. Revisa el archivo o contacta al administrador.") from exc
        finally:
            app.state.processing.release()
        return info

    @app.post("/api/excel/inspect")
    async def inspect_workbook(excel: UploadFile = File(...)):
        filename_safe(excel.filename, "xlsx")
        data = await read_upload(excel, MAX_PDF_BYTES)
        try:
            layout = await asyncio.to_thread(inspect_excel, data)
        except CertificateError as exc:
            raise HTTPException(422, str(exc)) from exc
        return {"sheetName": layout.sheet_name, "preserveFormat": layout.preserve_format,
                "needsBackground": not layout.preserve_format}

    app.mount("/static", StaticFiles(directory=STATIC), name="static")
    return app


app = create_app()
