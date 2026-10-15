"""
=============================================================================
MATRIZ MCE SIGRAMA - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia de Alto Rendimiento con Google Cloud Storage (Cloud Run)
- Cliente Singleton reutilizable
- Caché en memoria con TTL (300s) y control de concurrencia thread-safe
- Solo descarga archivos nuevos o modificados
=============================================================================
"""

import os
import time
import threading
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_DB = BASE_DIR / "base_matriz_mce.xlsx"
CATALOGO_FILE = BASE_DIR / "catalogos.json"
CARPETA_EVIDENCIAS = BASE_DIR / "evidencias"

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()

_STORAGE_CLIENT = None
_BUCKET_OBJ = None
_SYNC_LOCK = threading.Lock()
_LAST_SYNC_TS = 0.0
_GCS_READY = False
SYNC_TTL = 300.0  # 5 minutos


def get_storage_client():
    """Retorna una instancia singleton de storage.Client()."""
    global _STORAGE_CLIENT
    if _STORAGE_CLIENT is None:
        from google.cloud import storage
        _STORAGE_CLIENT = storage.Client()
    return _STORAGE_CLIENT


def _gcs_bucket():
    """Retorna una instancia singleton del bucket."""
    global _BUCKET_OBJ
    if not GCS_BUCKET:
        return None
    if _BUCKET_OBJ is None or _BUCKET_OBJ.name != GCS_BUCKET:
        _BUCKET_OBJ = get_storage_client().bucket(GCS_BUCKET)
    return _BUCKET_OBJ


def sync_from_gcs(force: bool = False) -> bool:
    """Descarga base_matriz_mce.xlsx, catalogos.json y evidencias desde GCS al arrancar o si expira el TTL."""
    global _GCS_READY, _LAST_SYNC_TS
    if not GCS_BUCKET:
        return False

    now = time.time()
    if not force and (now - _LAST_SYNC_TS) < SYNC_TTL and _GCS_READY:
        return True

    with _SYNC_LOCK:
        now = time.time()
        if not force and (now - _LAST_SYNC_TS) < SYNC_TTL and _GCS_READY:
            return True

        try:
            bucket = _gcs_bucket()
            if not bucket:
                return False

            count = 0
            CARPETA_EVIDENCIAS.mkdir(parents=True, exist_ok=True)
            for blob in bucket.list_blobs():
                if blob.name.endswith("/"):
                    continue
                if blob.name in ["data/base_matriz_mce.xlsx", "base_matriz_mce.xlsx"]:
                    if not ARCHIVO_DB.exists() or ARCHIVO_DB.stat().st_size != blob.size:
                        blob.download_to_filename(str(ARCHIVO_DB))
                        count += 1
                elif blob.name in ["data/catalogos.json", "catalogos.json"]:
                    if not CATALOGO_FILE.exists() or CATALOGO_FILE.stat().st_size != blob.size:
                        blob.download_to_filename(str(CATALOGO_FILE))
                        count += 1
                elif blob.name.startswith("evidencias/"):
                    rel_parts = blob.name.split("/")
                    local_path = BASE_DIR.joinpath(*rel_parts)
                    local_path.parent.mkdir(parents=True, exist_ok=True)
                    if not local_path.exists() or local_path.stat().st_size != blob.size:
                        blob.download_to_filename(str(local_path))
                        count += 1
            _GCS_READY = True
            _LAST_SYNC_TS = time.time()
            print(f"[GCS] Sincronizados exitosamente {count} archivos de Matriz MCE desde gs://{GCS_BUCKET}")
            return True
        except Exception as e:
            print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
            return False


def push_file_to_gcs(local_path):
    """Sube un archivo modificado o nuevo a GCS."""
    if not GCS_BUCKET:
        return False
    try:
        p = Path(local_path)
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        if not bucket:
            return False
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        if rel_path in ["base_matriz_mce.xlsx", "catalogos.json"]:
            bucket.blob(f"data/{rel_path}").upload_from_filename(str(p))
        print(f"[GCS] Archivo subido: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {local_path} a GCS: {e}")
        return False
