"""
=============================================================================
MATRIZ MCE SIGRAMA - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia con Google Cloud Storage (Cloud Run)
=============================================================================
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_DB = BASE_DIR / "base_matriz_mce.xlsx"
CATALOGO_FILE = BASE_DIR / "catalogos.json"
CARPETA_EVIDENCIAS = BASE_DIR / "evidencias"

GCS_BUCKET = os.environ.get("GCS_BUCKET", "").strip()
_GCS_READY = False

def _gcs_bucket():
    from google.cloud import storage
    return storage.Client().bucket(GCS_BUCKET)

def sync_from_gcs():
    """Descarga base_matriz_mce.xlsx, catalogos.json y evidencias desde GCS al arrancar."""
    global _GCS_READY
    if not GCS_BUCKET:
        return False
    try:
        bucket = _gcs_bucket()
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
        print(f"[GCS] Sincronizados exitosamente {count} archivos de Matriz MCE desde gs://{GCS_BUCKET}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
        return False

def push_file_to_gcs(local_path):
    """Sube un archivo modificado o nuevo a GCS."""
    if not (GCS_BUCKET and _GCS_READY):
        return False
    try:
        p = Path(local_path)
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        if rel_path in ["base_matriz_mce.xlsx", "catalogos.json"]:
            bucket.blob(f"data/{rel_path}").upload_from_filename(str(p))
        print(f"[GCS] Archivo subido: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {local_path} a GCS: {e}")
        return False
