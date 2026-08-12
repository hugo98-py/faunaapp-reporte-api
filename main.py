"""
API de generación de Reporte de Terreno (FaunaApp)
----------------------------------------------------
Recibe el JSON resumido de la página Reporte y devuelve el documento
generado en formato .docx o .pdf, según el parámetro `formato`.

Ejecutar localmente:
    uvicorn main:app --reload

Endpoint principal:
    POST /reportes/word?formato=docx   (o formato=pdf)
    Body: JSON con la forma de ReporteRequest (ver modelos abajo)
"""

import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import List, Literal, Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask

from generar_reporte import generar_reporte

app = FastAPI(
    title="API Reporte de Terreno - FaunaApp",
    description="Genera el Word/PDF del reporte de campaña a partir del resumen de la página Reporte.",
    version="1.0.0",
)


# ---------------------------------------------------------------------------
# Modelos de entrada (misma forma que REPORTE_JSON_SIMULADO en generar_reporte.py)
# ---------------------------------------------------------------------------

class EquipoProfesionalItem(BaseModel):
    nombre: str
    profesion: str
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
    payload: ReporteRequest,
    formato: Literal["docx", "pdf"] = Query("docx", description="Formato de descarga: docx o pdf"),
):
    """Genera el reporte y lo devuelve como archivo descargable."""

    # Carpeta temporal exclusiva por request, para evitar colisiones entre
    # descargas concurrentes en el mismo contenedor.
    tmp_dir = Path(tempfile.mkdtemp(prefix="reporte_"))
    nombre_base = f"Reporte_Terreno_{uuid.uuid4().hex[:8]}"
    docx_path = tmp_dir / f"{nombre_base}.docx"

    try:
        generar_reporte(payload.model_dump(), docx_path)
    except Exception as exc:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Error generando el documento: {exc}") from exc

    if formato == "docx":
        return FileResponse(
            path=docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{nombre_base}.docx",
            background=BackgroundTask(shutil.rmtree, tmp_dir, ignore_errors=True),
        )

    # formato == "pdf": convertir con LibreOffice headless
    pdf_path = tmp_dir / f"{nombre_base}.pdf"
    try:
        resultado = subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                "--convert-to", "pdf", "--outdir", str(tmp_dir), str(docx_path),
            ],
            capture_output=True, text=True, timeout=60,
        )
        if resultado.returncode != 0 or not pdf_path.exists():
            raise RuntimeError(resultado.stderr or "LibreOffice no generó el PDF.")
    except Exception as exc:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Error convirtiendo a PDF: {exc}") from exc

    return FileResponse(
        path=pdf_path,
        media_type="application/pdf",
        filename=f"{nombre_base}.pdf",
        background=BackgroundTask(shutil.rmtree, tmp_dir, ignore_errors=True),
    )
