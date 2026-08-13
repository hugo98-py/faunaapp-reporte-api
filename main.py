"""
API de generación de Reporte de Terreno (FaunaApp)
----------------------------------------------------
Recibe el JSON resumido de la página Reporte y genera el documento en
formato .docx o .pdf, según el parámetro `formato`.

Flujo en dos pasos (para que el cliente reciba una URL de descarga en
vez del archivo directo en la respuesta del POST):
    1. POST /reportes/word?formato=docx   -> {"url": "https://.../descargas/xxx.docx"}
    2. GET  esa url                       -> descarga el archivo (y se autoborra del servidor)

Ejecutar localmente:
    uvicorn main:app --reload
"""

import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import List, Literal, Optional

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from generar_reporte import generar_reporte

app = FastAPI(
    title="API Reporte de Terreno - FaunaApp",
    description="Genera el Word/PDF del reporte de campaña a partir del resumen de la página Reporte.",
    version="1.1.0",
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
DESCARGAS_DIR = Path(tempfile.gettempdir()) / "reportes_descargas"
DESCARGAS_DIR.mkdir(parents=True, exist_ok=True)

MEDIA_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pdf": "application/pdf",
}


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
                capture_output=True, text=True, timeout=60,
            )
            if resultado.returncode != 0 or not pdf_path.exists():
                raise RuntimeError(resultado.stderr or "LibreOffice no generó el PDF.")
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}") from exc
        finally:
            docx_path.unlink(missing_ok=True)  # ya no se necesita el .docx intermedio
        archivo_final = pdf_path

    # request.base_url ya trae el dominio correcto (local o el de Render),
    # sin necesidad de hardcodearlo.
    url_descarga = f"{str(request.base_url).rstrip('/')}/descargas/{archivo_final.name}"
    return {"url": url_descarga, "formato": formato, "nombreArchivo": archivo_final.name}


@app.get("/descargas/{nombre_archivo}")
def descargar_archivo(nombre_archivo: str):
    """Sirve el archivo generado y lo borra del servidor una vez entregado."""

    ruta = DESCARGAS_DIR / nombre_archivo
    # Evita path traversal (ej. ../../algo) validando que siga dentro de DESCARGAS_DIR
    if DESCARGAS_DIR not in ruta.resolve().parents or not ruta.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado o ya expiró.")

    extension = ruta.suffix.lstrip(".")
    media_type = MEDIA_TYPES.get(extension, "application/octet-stream")

    return FileResponse(
        path=ruta,
        media_type=media_type,
        filename=ruta.name,
        background=BackgroundTask(ruta.unlink, missing_ok=True),
    )
