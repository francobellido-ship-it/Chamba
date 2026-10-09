"""Local development storage and durable Google Drive publication."""

from dataclasses import dataclass
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import threading
import time
from uuid import UUID, uuid4

import httpx

logger = logging.getLogger(__name__)
DRIVE_API = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD = "https://www.googleapis.com/upload/drive/v3"
TOKEN_URL = "https://oauth2.googleapis.com/token"
FILE_ID = re.compile(r"[A-Za-z0-9_-]{10,200}")
MAX_DOWNLOAD = 64 * 1024 * 1024


class StorageError(Exception):
    def __init__(self, message: str, status: int = 503):
        self.status = status
        super().__init__(message)


@dataclass(frozen=True)
class Publication:
    verification_url: str
    drive_file_id: str | None = None


class LocalStorage:
    backend = "local"
    ready = True

    def __init__(self, directory: Path, quota: int):
        self.directory = directory
        self.quota = quota
        self.lock = threading.Lock()

    def start(self):
        self.directory.mkdir(parents=True, exist_ok=True)

    def close(self):
        pass

    def allocate(self, certificate_id: str, origin: str) -> Publication:
        return Publication(f"{origin}/certificados/{certificate_id}")

    def save(self, certificate_id: str, pdf: bytes, metadata: dict, publication: Publication):
        with self.lock:
            used = sum(path.stat().st_size for path in self.directory.glob("*/*") if path.is_file())
            if used + len(pdf) + 8192 > self.quota:
                raise StorageError("El almacenamiento está lleno. Contacta al administrador.", 507)
            directory = self.directory / str(UUID(certificate_id))
            if directory.exists():
                raise StorageError("El certificado ya existe; se conservó el registro anterior.", 409)
            try:
                directory.mkdir(mode=0o700)
                (directory / "certificate.pdf").write_bytes(pdf)
                temporary = directory / "metadata.tmp"
                temporary.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
                temporary.replace(directory / "metadata.json")
            except OSError:
                for name in ("certificate.pdf", "metadata.tmp", "metadata.json"):
                    (directory / name).unlink(missing_ok=True)
                if directory.exists():
                    directory.rmdir()
                raise StorageError("No se pudo guardar el certificado. Contacta al administrador.", 507) from None

    def metadata(self, certificate_id: str) -> dict:
        directory = self.directory / str(UUID(certificate_id))
        path = directory / "metadata.json"
        if not path.is_file() or not (directory / "certificate.pdf").is_file():
            raise StorageError("Certificado no encontrado.", 404)
        return json.loads(path.read_text(encoding="utf-8"))

    def local_pdf(self, certificate_id: str) -> Path:
        return self.directory / str(UUID(certificate_id)) / "certificate.pdf"


class DriveStorage:
    backend = "drive"

    def __init__(self, folder_id: str, client_id: str, client_secret: str, refresh_token: str,
                 client: httpx.Client | None = None):
        self.folder_id = folder_id.strip()
        self.client_id = client_id.strip()
        self.client_secret = client_secret.strip()
        self.refresh_token = refresh_token.strip()
        self.ready = bool(FILE_ID.fullmatch(self.folder_id) and self.client_id and self.client_secret and self.refresh_token)
        self.client = client or httpx.Client(timeout=httpx.Timeout(45, connect=10), follow_redirects=False)
        self.token = ""
        self.token_deadline = 0.0
        self.token_lock = threading.Lock()
        self.write_lock = threading.Lock()

    def start(self):
        # A missing configuration leaves the home page usable, with uploads disabled.
        pass

    def close(self):
        self.client.close()

    def _token(self) -> str:
        if not self.ready:
            raise StorageError("Falta completar la conexión con Google Drive. Contacta al administrador.")
        with self.token_lock:
            if time.monotonic() < self.token_deadline:
                return self.token
            try:
                response = self.client.post(TOKEN_URL, data={
                    "grant_type": "refresh_token", "client_id": self.client_id,
                    "client_secret": self.client_secret, "refresh_token": self.refresh_token,
                })
                if response.status_code in {400, 401, 403}:
                    raise StorageError("La autorización de Google Drive no es válida o fue revocada. Contacta al administrador.")
                response.raise_for_status()
                result = response.json()
                token = result["access_token"]
                expires = int(result.get("expires_in", 3600))
                if not isinstance(token, str) or not token or expires <= 60:
                    raise ValueError
            except (httpx.HTTPError, ValueError, KeyError, TypeError):
                raise StorageError("No se pudo renovar la conexión con Google Drive. Inténtalo de nuevo.") from None
            self.token = token
            self.token_deadline = time.monotonic() + expires - 60
            return token

    def _request(self, method: str, path: str, *, upload: bool = False, **kwargs) -> httpx.Response:
        headers = {"Authorization": f"Bearer {self._token()}", **kwargs.pop("headers", {})}
        try:
            response = self.client.request(method, (DRIVE_UPLOAD if upload else DRIVE_API) + path,
                                           headers=headers, **kwargs)
        except httpx.HTTPError:
            raise StorageError("No se pudo conectar con Google Drive. Inténtalo de nuevo.") from None
        if response.status_code == 404:
            raise StorageError("El archivo o la carpeta no está disponible en Google Drive.", 404)
        if response.status_code == 403:
            raise StorageError("Google Drive rechazó la operación. Revisa los permisos y el espacio de la cuenta.")
        if response.status_code == 401:
            self.token_deadline = 0
            raise StorageError("Google Drive requiere renovar su autorización. Inténtalo de nuevo.")
        if response.status_code < 200 or response.status_code >= 300:
            raise StorageError("Google Drive no pudo completar la operación. Inténtalo de nuevo.")
        return response

    def _json(self, response: httpx.Response) -> dict:
        try:
            result = response.json()
            if not isinstance(result, dict):
                raise ValueError
            return result
        except ValueError:
            raise StorageError("Google Drive devolvió una respuesta no válida.") from None

    def _root(self):
        folder = self._json(self._request("GET", f"/files/{self.folder_id}", params={
            "fields": "id,mimeType,trashed,capabilities(canAddChildren)", "supportsAllDrives": "true"}))
        if folder.get("mimeType") != "application/vnd.google-apps.folder" or folder.get("trashed"):
            raise StorageError("La carpeta configurada en Google Drive no está disponible.")
        if folder.get("capabilities", {}).get("canAddChildren") is not True:
            raise StorageError("La cuenta conectada no puede guardar archivos en la carpeta de Google Drive.")
        permissions = self._json(self._request("GET", f"/files/{self.folder_id}/permissions", params={
            "fields": "permissions(type),nextPageToken", "pageSize": 100, "supportsAllDrives": "true"}))
        if permissions.get("nextPageToken") or any(row.get("type") in {"anyone", "domain"}
                                                   for row in permissions.get("permissions", [])):
            raise StorageError("Usa una carpeta de Drive restringida a personas concretas; solo se publicará cada PDF final.")

    def _list(self, parent: str, extra: str) -> list[dict]:
        result = self._json(self._request("GET", "/files", params={
            "q": f"'{parent}' in parents and trashed = false and ({extra})",
            "fields": "files(id,name,mimeType),nextPageToken", "pageSize": 100,
            "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}))
        if result.get("nextPageToken"):
            raise StorageError("Hay registros duplicados o inesperados en la carpeta de Drive.")
        rows = result.get("files", [])
        if not isinstance(rows, list) or any(not isinstance(row, dict) or not FILE_ID.fullmatch(str(row.get("id", ""))) for row in rows):
            raise StorageError("Google Drive devolvió un listado no válido.")
        return rows

    def _record_folder(self, certificate_id: str) -> str:
        identifier = str(UUID(certificate_id))
        rows = self._list(self.folder_id, "mimeType = 'application/vnd.google-apps.folder' and "
                          f"appProperties has {{ key='qcp_certificate_id' and value='{identifier}' }}")
        if not rows:
            raise StorageError("Certificado no encontrado.", 404)
        if len(rows) != 1:
            raise StorageError("Hay registros duplicados para este certificado. Contacta al administrador.")
        return self._id(rows[0]["id"])

    def _id(self, value: str) -> str:
        if not isinstance(value, str) or not FILE_ID.fullmatch(value):
            raise StorageError("Google Drive devolvió un identificador no válido.")
        return value

    def allocate(self, certificate_id: str, origin: str) -> Publication:
        self._root()
        result = self._json(self._request("GET", "/files/generateIds", params={"count": 1, "space": "drive"}))
        try:
            file_id = self._id(result["ids"][0])
        except (KeyError, IndexError, TypeError):
            raise StorageError("No se pudo reservar el enlace del certificado en Google Drive.") from None
        return Publication(f"https://drive.google.com/file/d/{file_id}/view", file_id)

    def _upload(self, metadata: dict, data: bytes, mime: str) -> str:
        boundary = "qcp_" + uuid4().hex
        body = (f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n".encode()
                + json.dumps(metadata, ensure_ascii=False).encode() + b"\r\n"
                + f"--{boundary}\r\nContent-Type: {mime}\r\n\r\n".encode()
                + data + f"\r\n--{boundary}--\r\n".encode())
        result = self._json(self._request("POST", "/files", upload=True,
            params={"uploadType": "multipart", "supportsAllDrives": "true", "fields": "id"},
            headers={"Content-Type": f"multipart/related; boundary={boundary}"}, content=body))
        return self._id(result.get("id"))

    def save(self, certificate_id: str, pdf: bytes, metadata: dict, publication: Publication):
        identifier = str(UUID(certificate_id))
        file_id = self._id(publication.drive_file_id)
        with self.write_lock:
            existing = self._list(self.folder_id, "mimeType = 'application/vnd.google-apps.folder' and "
                                 f"appProperties has {{ key='qcp_certificate_id' and value='{identifier}' }}")
            if existing:
                raise StorageError("El certificado ya existe; se conservó el registro anterior.", 409)
            folder = self._json(self._request("POST", "/files", params={"supportsAllDrives": "true", "fields": "id"}, json={
                "name": f"{metadata['filename'][:-4]} · {identifier[:8]}",
                "mimeType": "application/vnd.google-apps.folder", "parents": [self.folder_id],
                "appProperties": {"qcp_certificate_id": identifier},
            }))
            record_folder = self._id(folder.get("id"))
            try:
                actual_id = self._upload({"id": file_id, "name": metadata["filename"], "mimeType": "application/pdf",
                                         "parents": [record_folder]}, pdf, "application/pdf")
                if actual_id != file_id:
                    raise StorageError("Google Drive no conservó el enlace reservado del certificado.")
                # Only the final PDF is public. Folders and JSON records stay private.
                self._request("POST", f"/files/{file_id}/permissions", params={"supportsAllDrives": "true"},
                              json={"type": "anyone", "role": "reader", "allowFileDiscovery": False})
                metadata["storageBackend"] = "drive"
                metadata["driveFileId"] = file_id
                metadata["driveUrl"] = f"https://drive.google.com/file/d/{file_id}/view"
                self._upload({"name": "metadata.json", "mimeType": "application/json", "parents": [record_folder]},
                             json.dumps(metadata, ensure_ascii=False).encode(), "application/json")
            except Exception:
                try:
                    self._request("DELETE", f"/files/{record_folder}", params={"supportsAllDrives": "true"})
                except StorageError:
                    logger.warning("No se pudo limpiar un registro de publicación incompleta en Drive.")
                raise

    def metadata(self, certificate_id: str) -> dict:
        folder = self._record_folder(certificate_id)
        records = self._list(folder, "name = 'metadata.json' and mimeType = 'application/json'")
        if len(records) != 1:
            raise StorageError("El registro de este certificado está incompleto.", 404)
        response = self._request("GET", f"/files/{self._id(records[0]['id'])}", params={"alt": "media"})
        if len(response.content) > 64 * 1024:
            raise StorageError("El registro del certificado supera el tamaño permitido.")
        info = self._json(response)
        if (info.get("id") != str(UUID(certificate_id))
                or not isinstance(info.get("driveFileId"), str) or not FILE_ID.fullmatch(info["driveFileId"])
                or not isinstance(info.get("filename"), str) or not re.fullmatch(r"[a-f0-9]{64}", str(info.get("sha256", "")))):
            raise StorageError("El registro del certificado no es válido.")
        return info

    def download(self, metadata: dict) -> bytes:
        file_id = self._id(metadata.get("driveFileId"))
        headers = {"Authorization": f"Bearer {self._token()}"}
        try:
            with self.client.stream("GET", f"{DRIVE_API}/files/{file_id}", params={"alt": "media"}, headers=headers) as response:
                if response.status_code == 404:
                    raise StorageError("Certificado no encontrado.", 404)
                if response.status_code != 200:
                    raise StorageError("Google Drive no pudo entregar el certificado. Inténtalo de nuevo.")
                chunks = []
                size = 0
                for chunk in response.iter_bytes():
                    size += len(chunk)
                    if size > MAX_DOWNLOAD:
                        raise StorageError("El PDF almacenado supera el tamaño permitido.")
                    chunks.append(chunk)
                pdf = b"".join(chunks)
        except httpx.HTTPError:
            raise StorageError("No se pudo descargar el certificado desde Google Drive.") from None
        if hashlib.sha256(pdf).hexdigest() != metadata.get("sha256"):
            raise StorageError("El PDF almacenado cambió y no coincide con el registro del certificado.", 409)
        return pdf


def storage_from_environment(directory: Path, quota: int) -> LocalStorage | DriveStorage:
    backend = os.environ.get("CERTIFICATE_STORAGE_BACKEND", "local").strip().lower()
    if backend == "local":
        return LocalStorage(directory, quota)
    if backend != "drive":
        raise ValueError("CERTIFICATE_STORAGE_BACKEND debe ser local o drive.")
    return DriveStorage(*(os.environ.get(name, "") for name in (
        "GOOGLE_DRIVE_FOLDER_ID", "GOOGLE_DRIVE_CLIENT_ID", "GOOGLE_DRIVE_CLIENT_SECRET", "GOOGLE_DRIVE_REFRESH_TOKEN")))
