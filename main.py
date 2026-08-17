
"""
API de generación de Reporte de Terreno (FaunaApp)
----------------------------------------------------
Recibe el JSON resumido de la página Reporte y genera el documento en
formato .docx o .pdf, según el parámetro `formato`.

Flujo en dos pasos (para que el cliente reciba una URL de descarga en
vez del archivo directo en la respuesta del POST):
    1. POST /reportes/word?formato=docx   -> {"url": "https://.../descargas/xxx.docx"}
    2. GET  esa url                       -> descarga el archivo

IMPORTANTE (esquema http vs https):
    Render termina el TLS en su proxy y le habla a uvicorn por HTTP plano,
    así que `request.base_url` devuelve http:// salvo que uvicorn confíe en
    el header X-Forwarded-Proto. Si la URL de descarga sale con http://, el
    navegador BLOQUEA la descarga en silencio (mixed content download), que
    es justo el síntoma de "se abre la pestaña un instante y no baja nada".
    Por eso acá la URL se arma leyendo los headers del proxy directamente,
    y además el Dockerfile define FORWARDED_ALLOW_IPS=* + --proxy-headers.

Ejecutar localmente:
    uvicorn main:app --reload
"""

import re
import subprocess
import tempfile
import time
import uuid
from pathlib import Path
from typing import List, Literal

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from generar_reporte import generar_reporte

app = FastAPI(
    title="API Reporte de Terreno - FaunaApp",
    description="Genera el Word/PDF del reporte de campaña a partir del resumen de la página Reporte.",
    version="1.2.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

# Carpeta donde quedan los archivos ya generados, esperando a que el cliente
# haga el GET de descarga. Vive mientras el contenedor esté corriendo (se
# reinicia con cada deploy, lo cual está bien: son archivos de paso, no datos
# a conservar).
DESCARGAS_DIR = (Path(tempfile.gettempdir()) / "reportes_descargas").resolve()
DESCARGAS_DIR.mkdir(parents=True, exist_ok=True)

# Los archivos se borran por antigüedad, NO en el primer GET. Si se borraran
# al servirlos, un reintento del navegador (o una descarga bloqueada que
# igual alcanzó a pegarle al server) dejaría el link quemado con un 404.
VIDA_ARCHIVO_SEGUNDOS = 30 * 60

MEDIA_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pdf": "application/pdf",
}

# Solo nombres tipo Reporte_Terreno_ab12cd34.docx — corta cualquier intento
# de path traversal antes de tocar el disco.
NOMBRE_VALIDO = re.compile(r"^[A-Za-z0-9_\-]+\.(docx|pdf)$")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def limpiar_antiguos() -> None:
    """Borra los archivos de descarga que ya pasaron su tiempo de vida."""
    limite = time.time() - VIDA_ARCHIVO_SEGUNDOS
    for archivo in DESCARGAS_DIR.glob("*"):
        try:
            if archivo.is_file() and archivo.stat().st_mtime < limite:
                archivo.unlink(missing_ok=True)
        except OSError:
            pass  # si otro request lo borró justo ahora, da lo mismo


def base_url_publica(request: Request) -> str:
    """
    Reconstruye la URL pública real (la que ve el navegador), respetando el
    proxy de Render. `request.base_url` no sirve por sí solo: si uvicorn no
    confía en el proxy, devuelve http:// y el navegador bloquea la descarga.
    """
    # X-Forwarded-Proto puede venir como "https, http" si hay varios saltos;
    # el primero es el del cliente.
    proto = request.headers.get("x-forwarded-proto", "").split(",")[0].strip()
    if not proto:
        proto = request.url.scheme

    host = request.headers.get("x-forwarded-host", "").split(",")[0].strip()
    if not host:
        host = request.headers.get("host", "").strip() or request.url.netloc

    # Red de seguridad: cualquier host que no sea local va sí o sí por https.
    if not host.startswith(("localhost", "127.0.0.1", "0.0.0.0")):
        proto = "https"

    return f"{proto}://{host}"


# ---------------------------------------------------------------------------
# Modelos de entrada (misma forma que REPORTE_JSON_SIMULADO en generar_reporte.py)
# ---------------------------------------------------------------------------

class EquipoProfesionalItem(BaseModel):
    nombre: str
    profesion: str = ""
    jefeTerreno: bool = False


class Metadata(BaseModel):
    nombreProyecto: str
    epoca: str
    anio: str
    mesAnio: str
    fechaCampana: str
    equipoProfesional: List[EquipoProfesionalItem]
    equiposUtilizados: str
    tipoCampana: str = ""
    componente: str = ""


class SingularidadTabla2Item(BaseModel):
    clase: str
    nombreCientifico: str
    nombreComun: str
    categoriaConservacion: str
    movilidad: str
    origen: str


class SingularidadTabla3Item(BaseModel):
    clase: str
    nombreCientifico: str
    nameest: str
    abundancia: int


class ReporteRequest(BaseModel):
    metadata: Metadata
    singularidadesTabla2: List[SingularidadTabla2Item]
    singularidadesTabla3: List[SingularidadTabla3Item]


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/")
def health_check():
    """Endpoint simple para el health check de Render."""
    return {"status": "ok"}


@app.post("/reportes/word")
def generar_documento(
    request: Request,
    payload: ReporteRequest,
    formato: Literal["docx", "pdf"] = Query("docx", description="Formato de descarga: docx o pdf"),
):
    """Genera el reporte y devuelve la URL para descargarlo (no el archivo directo)."""

    limpiar_antiguos()

    nombre_base = f"Reporte_Terreno_{uuid.uuid4().hex[:8]}"
    docx_path = DESCARGAS_DIR / f"{nombre_base}.docx"

    try:
        generar_reporte(payload.model_dump(), docx_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error generando el documento: {exc}") from exc

    archivo_final = docx_path

    if formato == "pdf":
        pdf_path = DESCARGAS_DIR / f"{nombre_base}.pdf"
        try:
            resultado = subprocess.run(
                [
                    "soffice", "--headless", "--norestore",
                    "--convert-to", "pdf", "--outdir", str(DESCARGAS_DIR), str(docx_path),
                ],
                capture_output=True, text=True, timeout=120,
            )
            if resultado.returncode != 0 or not pdf_path.exists():
                raise RuntimeError(resultado.stderr or "LibreOffice no generó el PDF.")
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}") from exc
        finally:
            docx_path.unlink(missing_ok=True)  # ya no se necesita el .docx intermedio
        archivo_final = pdf_path

    url_descarga = f"{base_url_publica(request)}/descargas/{archivo_final.name}"
    return {"url": url_descarga, "formato": formato, "nombreArchivo": archivo_final.name}


@app.get("/descargas/{nombre_archivo}")
def descargar_archivo(nombre_archivo: str):
    """Sirve el archivo generado. El borrado se hace por antigüedad, no acá."""

    if not NOMBRE_VALIDO.match(nombre_archivo):
        raise HTTPException(status_code=404, detail="Archivo no encontrado o ya expiró.")

    ruta = (DESCARGAS_DIR / nombre_archivo).resolve()
    if ruta.parent != DESCARGAS_DIR or not ruta.is_file():
        raise HTTPException(status_code=404, detail="Archivo no encontrado o ya expiró.")

    extension = ruta.suffix.lstrip(".")
    media_type = MEDIA_TYPES.get(extension, "application/octet-stream")

    return FileResponse(
        path=ruta,
        media_type=media_type,
        filename=ruta.name,
        headers={"Cache-Control": "no-store"},
    )
