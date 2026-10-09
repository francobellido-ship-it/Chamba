"""Read-only check of the configured Google account and destination folder."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.storage import DriveStorage, StorageError, storage_from_environment


def main():
    storage = storage_from_environment(Path("data/certificates"), 1024 * 1024 * 1024)
    try:
        if not isinstance(storage, DriveStorage):
            raise SystemExit("Selecciona CERTIFICATE_STORAGE_BACKEND=drive para comprobar la conexión.")
        if not storage.ready:
            raise SystemExit("Faltan las variables de Google Drive. Configúralas sin compartir secretos en el chat.")
        storage._root()
        print("Conexión autorizada: carpeta accesible, restringida y con permiso para guardar archivos.")
        print("Comprobación de solo lectura; no se creó ni publicó ningún certificado.")
    except StorageError as error:
        raise SystemExit(str(error)) from None
    finally:
        storage.close()


if __name__ == "__main__":
    main()
