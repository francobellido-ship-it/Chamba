from contextlib import asynccontextmanager
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timezone
import asyncio
import hashlib
import ipaddress
import json
import logging
import multiprocessing
import os
from pathlib import Path
import re
import secrets
from urllib.parse import urlsplit
from uuid import UUID, uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException

from app.pdf import CertificateError, prepare_pdf

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
    """Reject unauthorized uploads before parsing files; bound streamed request bodies."""

    def __init__(self, app, access_key: str, uploads_enabled: bool):
        self.app = app
        self.access_key = access_key
        self.uploads_enabled = uploads_enabled

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)
        is_upload = scope["method"] == "POST" and scope["path"] == "/api/certificates"
        if not is_upload:
            return await self.app(scope, receive, send)
        if not self.uploads_enabled:
            return await JSONResponse(
                {"detail": "La preparación de certificados todavía no está habilitada. Contacta al administrador."},
                status_code=503,
            )(scope, receive, send)
        headers = dict(scope["headers"])
        if self.access_key:
            supplied = headers.get(b"x-portal-key", b"")
            if not secrets.compare_digest(supplied, self.access_key.encode()):
                return await JSONResponse({"detail": "La clave de acceso no es correcta."}, status_code=401)(scope, receive, send)
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


def filename_safe(name: str | None) -> str:
    name = (name or "CERTIFICADO.pdf").replace("\\", "/").split("/")[-1]
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip()
    if not name.lower().endswith(".pdf"):
        raise HTTPException(422, "El archivo debe tener extensión .pdf.")
    return name[:180 - 4].removesuffix(".pdf").removesuffix(".PDF") + ".pdf" if len(name) > 180 else name


def create_app(storage_dir: Path | None = None, public_url: str | None = None,
               access_key: str | None = None) -> FastAPI:
    storage = storage_dir or Path(os.environ.get("CERTIFICATE_STORAGE_DIR", "data/certificates"))
    origin = (public_url or os.environ.get("PUBLIC_BASE_URL") or os.environ.get("RENDER_EXTERNAL_URL")
              or "http://127.0.0.1:8000").rstrip("/")
    parsed = urlsplit(origin)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.path or parsed.query or parsed.fragment or parsed.username or parsed.password:
        raise ValueError("PUBLIC_BASE_URL debe ser una URL http(s) sin ruta, credenciales, consulta ni fragmento.")
    key = access_key if access_key is not None else os.environ.get("PORTAL_ACCESS_KEY", "")
    is_local = local_url(origin)
    if not is_local and parsed.scheme != "https":
        raise ValueError("PUBLIC_BASE_URL debe usar HTTPS para publicar el portal.")
    uploads_enabled = is_local or len(key) >= 16
    quota = int(os.environ.get("MAX_STORAGE_MB", "1024")) * 1024 * 1024

    @asynccontextmanager
    async def lifespan(app):
        storage.mkdir(parents=True, exist_ok=True)
        if not uploads_enabled:
            logger.warning(
                "La página pública está disponible, pero la preparación de certificados está deshabilitada. "
                "Configura PORTAL_ACCESS_KEY con al menos 16 caracteres y vuelve a desplegar."
            )
        app.state.executor = ProcessPoolExecutor(max_workers=2, mp_context=multiprocessing.get_context("spawn"))
        app.state.processing = asyncio.Semaphore(2)
        app.state.storage_lock = asyncio.Lock()
        yield
        app.state.executor.shutdown(wait=True, cancel_futures=True)

    app = FastAPI(title="Portal de certificados QCP", lifespan=lifespan,
                  docs_url=None, redoc_url=None, openapi_url=None)
    app.add_middleware(RequestGuard, access_key=key, uploads_enabled=uploads_enabled)

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
        return {"requiresAccessKey": bool(key) or not is_local, "uploadsEnabled": uploads_enabled, "localMode": is_local,
                "maxPdfMB": 15, "maxSignatureMB": 2, "maxPages": 100}

    @app.get("/")
    async def home():
        return FileResponse(STATIC / "index.html")

    @app.get("/certificados/{certificate_id}")
    async def certificate_page(certificate_id: UUID):
        if not (storage / str(certificate_id) / "metadata.json").is_file():
            raise HTTPException(404, "Certificado no encontrado.")
        return FileResponse(STATIC / "verify.html", headers={"Cache-Control": "no-store"})

    def metadata(certificate_id: UUID) -> dict:
        path = storage / str(certificate_id) / "metadata.json"
        pdf = storage / str(certificate_id) / "certificate.pdf"
        if not path.is_file() or not pdf.is_file():
            raise HTTPException(404, "Certificado no encontrado.")
        return json.loads(path.read_text(encoding="utf-8"))

    @app.get("/api/certificates/{certificate_id}")
    async def details(certificate_id: UUID):
        return metadata(certificate_id)

    @app.get("/api/certificates/{certificate_id}/download")
    async def download(certificate_id: UUID):
        info = metadata(certificate_id)
        return FileResponse(storage / str(certificate_id) / "certificate.pdf", media_type="application/pdf",
                            filename=info["filename"], headers={"Cache-Control": "no-store"})

    @app.post("/api/certificates", status_code=201)
    async def create_certificate(pdf: UploadFile = File(...), codigoCorrecto: str = Form(""),
                                 firma: UploadFile | None = File(None)):
        filename = filename_safe(pdf.filename)
        if len(codigoCorrecto) > 40:
            raise HTTPException(422, "El código debe tener el formato QCP-0000-0000.")
        data = await read_upload(pdf, MAX_PDF_BYTES)
        signature = await read_upload(firma, MAX_SIGNATURE_BYTES) if firma else None
        certificate_id = str(uuid4())
        verification_url = f"{origin}/certificados/{certificate_id}"
        try:
            await asyncio.wait_for(app.state.processing.acquire(), timeout=0.1)
        except TimeoutError:
            raise HTTPException(429, "El portal está procesando otros archivos. Inténtalo nuevamente en unos segundos.")
        try:
            result = await asyncio.get_running_loop().run_in_executor(
                app.state.executor, prepare_pdf, data, codigoCorrecto, verification_url, signature)
        except CertificateError as exc:
            raise HTTPException(422, str(exc)) from exc
        except Exception as exc:
            logger.exception("Error al procesar un certificado")
            raise HTTPException(500, "No se pudo procesar el PDF. Revisa el archivo o contacta al administrador.") from exc
        finally:
            app.state.processing.release()
        info = {"id": certificate_id, "filename": filename,
                "createdAt": datetime.now(timezone.utc).isoformat(),
                "detectedCodes": result.detected_codes, "finalCode": result.final_code,
                "replacements": result.replacements, "warnings": result.warnings,
                "stampPages": result.stamp_pages, "signatureIncluded": result.signature_included,
                "sha256": hashlib.sha256(result.data).hexdigest(),
                "verificationUrl": verification_url,
                "downloadUrl": f"/api/certificates/{certificate_id}/download"}
        async with app.state.storage_lock:
            used = sum(path.stat().st_size for path in storage.glob("*/*") if path.is_file())
            if used + len(result.data) + 8192 > quota:
                raise HTTPException(507, "El almacenamiento está lleno. Contacta al administrador.")
            directory = storage / certificate_id
            try:
                directory.mkdir(mode=0o700)
                (directory / "certificate.pdf").write_bytes(result.data)
                temp_metadata = directory / "metadata.tmp"
                temp_metadata.write_text(json.dumps(info, ensure_ascii=False), encoding="utf-8")
                temp_metadata.replace(directory / "metadata.json")
            except OSError as exc:
                for name in ("certificate.pdf", "metadata.tmp", "metadata.json"):
                    (directory / name).unlink(missing_ok=True)
                if directory.exists():
                    directory.rmdir()
                logger.exception("No se pudo guardar el certificado")
                raise HTTPException(507, "No se pudo guardar el certificado. Contacta al administrador.") from exc
        return info

    app.mount("/static", StaticFiles(directory=STATIC), name="static")
    return app


app = create_app()
