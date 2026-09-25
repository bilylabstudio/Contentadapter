"""
Adaptador de presentaciones IMSED -- app web stateless.

Sube 1 o varias presentaciones antiguas (.pptx / .pdf) y devuelve la(s)
versión(es) adaptadas al diseño de la plantilla madre IMSED. No se guarda
nada en base de datos ni en disco de forma persistente: cada request
procesa sus archivos en memoria / un directorio temporal que se borra al
terminar la respuesta.

La API key de Gemini la manda el navegador en cada llamada (campo del
formulario) -- no se guarda en el servidor ni en ningún sitio.
"""
from __future__ import annotations

import asyncio
import io
import os
import tempfile
import zipfile
from pathlib import Path

import httpx
from flask import Flask, request, send_file, jsonify, Response

from engine.build import build_deck
from engine.extract import extract
from engine.mapper import map_deck

BASE_DIR = Path(__file__).parent
MASTER_TEMPLATE = BASE_DIR / "assets" / "master_template.pptx"

# En despliegues sin filesystem persistente (Vercel: solo /tmp es
# escribible y el repo no incluye el binario) descargamos la plantilla
# madre bajo demanda desde su URL pública de exportación de Google Slides
# y la cacheamos en /tmp para las siguientes invocaciones "calientes".
MASTER_TEMPLATE_URL = (
    "https://docs.google.com/presentation/d/"
    "1QO3ipklYEpNgGnbmBZ2wJm_ivOnv3pHvCbqV7Wezlbk/export/pptx"
)
MASTER_TEMPLATE_CACHE = Path(tempfile.gettempdir()) / "imsed_master_template.pptx"

MAX_FILES = 25
MAX_FILE_MB = 60
MAX_CONTENT_LENGTH = MAX_FILES * MAX_FILE_MB * 1024 * 1024

app = Flask(__name__, static_folder=str(BASE_DIR / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.after_request
def add_cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "*"
    return resp


@app.get("/")
def index():
    return (BASE_DIR / "static" / "index.html").read_text(encoding="utf-8")


@app.get("/health")
def health():
    return jsonify({
        "ok": True,
        "master_template_bundled": MASTER_TEMPLATE.exists(),
        "master_template_cached": MASTER_TEMPLATE_CACHE.exists(),
    })


class AppError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


@app.errorhandler(AppError)
def handle_app_error(e: AppError):
    return jsonify({"detail": e.message}), e.status


def _ensure_master_template() -> Path:
    """Devuelve la ruta a la plantilla madre, descargándola si hace falta.

    En Docker/EasyPanel el binario viene incluido en assets/ (se usa tal
    cual). En Vercel el repo no lo incluye -- se descarga la primera vez
    desde la URL pública de exportación de Google Slides y se cachea en
    /tmp para invocaciones posteriores del mismo contenedor.
    """
    if MASTER_TEMPLATE.exists():
        return MASTER_TEMPLATE

    if MASTER_TEMPLATE_CACHE.exists() and MASTER_TEMPLATE_CACHE.stat().st_size > 50_000:
        return MASTER_TEMPLATE_CACHE

    try:
        resp = httpx.get(MASTER_TEMPLATE_URL, follow_redirects=True, timeout=30.0)
        resp.raise_for_status()
    except Exception as e:
        raise AppError(500, f"No se pudo descargar la plantilla madre: {e}")

    if len(resp.content) < 50_000:
        raise AppError(500, "La plantilla madre descargada es demasiado pequeña; revisa el enlace público de Google Slides")

    MASTER_TEMPLATE_CACHE.write_bytes(resp.content)
    return MASTER_TEMPLATE_CACHE


async def _process_one(filename: str, raw_bytes: bytes, api_key: str, master_path: Path) -> bytes:
    slides = extract(filename, io.BytesIO(raw_bytes))
    if not slides:
        raise AppError(400, f"No se pudo extraer contenido de {filename}")

    specs_raw = await map_deck(slides, filename, api_key)

    errors = [s for s in specs_raw if s and "_error" in s]
    if errors and len(errors) == len(specs_raw):
        raise AppError(502, f"Gemini no respondió correctamente para {filename}: {errors[0]['_error']}")

    specs = [s for s in specs_raw if s and "_error" not in s]
    if not specs:
        raise AppError(400, f"No se generó ninguna diapositiva adaptada para {filename}")

    with tempfile.TemporaryDirectory() as tmp:
        out_path = Path(tmp) / "output.pptx"
        build_deck(str(master_path), str(out_path), specs)
        return out_path.read_bytes()


async def _process_batch(files_data: list[tuple[str, bytes]], api_key: str, master_path: Path):
    results: list[tuple[str, bytes]] = []
    errors: list[str] = []
    for filename, raw in files_data:
        try:
            out_bytes = await _process_one(filename, raw, api_key, master_path)
            out_name = Path(filename).stem + "_IMSED.pptx"
            results.append((out_name, out_bytes))
        except AppError as e:
            errors.append(f"{filename}: {e.message}")
        except Exception as e:
            errors.append(f"{filename}: error inesperado ({e})")
    return results, errors


@app.post("/transform")
def transform():
    api_key = request.form.get("gemini_api_key", "").strip()
    if not api_key:
        raise AppError(400, "Falta la API key de Gemini")

    uploaded = request.files.getlist("files")
    if not uploaded:
        raise AppError(400, "No se ha subido ningún archivo")
    if len(uploaded) > MAX_FILES:
        raise AppError(400, f"Máximo {MAX_FILES} archivos por lote")

    master_path = _ensure_master_template()

    files_data = []
    skip_errors = []
    for f in uploaded:
        raw = f.read()
        if len(raw) > MAX_FILE_MB * 1024 * 1024:
            skip_errors.append(f"{f.filename}: supera los {MAX_FILE_MB}MB, se omite")
            continue
        files_data.append((f.filename, raw))

    results, errors = asyncio.run(_process_batch(files_data, api_key, master_path))
    errors = skip_errors + errors

    if not results:
        raise AppError(422, "No se pudo adaptar ningún archivo. " + " | ".join(errors))

    if len(results) == 1 and not errors:
        name, data = results[0]
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=name,
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in results:
            zf.writestr(name, data)
        if errors:
            zf.writestr("errores.txt", "\n".join(errors))
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="presentaciones_adaptadas.zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "80"))
    app.run(host="0.0.0.0", port=port, threaded=True)
