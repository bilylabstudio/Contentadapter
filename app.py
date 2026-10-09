"""
Adaptador de presentaciones IMSED -- app web stateless.

Sube 1 o varias presentaciones antiguas (.pptx / .pdf) y devuelve la(s)
versión(es) adaptadas al diseño de la plantilla madre IMSED. No se guarda
nada en base de datos ni en disco de forma persistente: cada request
procesa sus archivos en memoria / un directorio temporal que se borra al
terminar la respuesta.

La API key de Gemini se configura UNA VEZ como variable de entorno
(GEMINI_API_KEY) en el propio despliegue (Vercel / EasyPanel) -- el
formulario ya no la pide en cada uso. Sigue sin persistirse en ningún
sitio: solo vive en la config del entorno del servidor, nunca en disco
ni en base de datos.
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
from engine.masters import forms_url_for, list_masters, masters_without_url, resolve_master

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

# Se configura una sola vez como variable de entorno del despliegue
# (Settings -> Environment Variables en Vercel, o env del contenedor en
# EasyPanel). El formulario ya no pide la key en cada transformación;
# si por lo que sea se manda igualmente en el form (compatibilidad hacia
# atrás), esa tiene prioridad sobre la de entorno.
GEMINI_API_KEY_ENV = os.environ.get("GEMINI_API_KEY", "").strip()

MAX_FILES = 25
MAX_FILE_MB = 60
# Límite total por subida (todos los archivos del lote juntos). En Vercel
# esto es irrelevante -- la plataforma corta en 4.5MB antes de llegar aquí,
# por eso ese despliegue solo sirve para lotes pequeños -- pero en
# EasyPanel/Docker (sin ese límite de plataforma) esto es lo único que
# protege al servidor de una subida descontrolada.
MAX_TOTAL_MB = 500
MAX_CONTENT_LENGTH = MAX_TOTAL_MB * 1024 * 1024

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
        "gemini_api_key_configured": bool(GEMINI_API_KEY_ENV),
        "masters": len(list_masters()),
        "masters_without_forms_url": masters_without_url(),
    })


@app.get("/masters")
def masters():
    """Másters disponibles (para el desplegable de la web y para n8n)."""
    return jsonify(list_masters())


class AppError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


@app.errorhandler(AppError)
def handle_app_error(e: AppError):
    return jsonify({"detail": e.message}), e.status


@app.errorhandler(413)
def handle_too_large(e):
    return jsonify({
        "detail": f"El lote supera el límite de {MAX_TOTAL_MB}MB por subida. "
                   "Divide los archivos en varias subidas más pequeñas."
    }), 413


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


async def _process_one(filename: str, raw_bytes: bytes, api_key: str, master_path: Path, qr_url: str) -> bytes:
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
        build_deck(str(master_path), str(out_path), specs, qr_url=qr_url)
        return out_path.read_bytes()


async def _process_batch(files_data: list[tuple[str, bytes]], api_key: str, master_path: Path, qr_url: str):
    results: list[tuple[str, bytes]] = []
    errors: list[str] = []
    for filename, raw in files_data:
        try:
            out_bytes = await _process_one(filename, raw, api_key, master_path, qr_url)
            out_name = Path(filename).stem + "_IMSED.pptx"
            results.append((out_name, out_bytes))
        except AppError as e:
            errors.append(f"{filename}: {e.message}")
        except Exception as e:
            errors.append(f"{filename}: error inesperado ({e})")
    return results, errors


@app.post("/transform")
def transform():
    api_key = request.form.get("gemini_api_key", "").strip() or GEMINI_API_KEY_ENV
    if not api_key:
        raise AppError(
            400,
            "Falta la API key de Gemini: configura la variable de entorno "
            "GEMINI_API_KEY en el despliegue (Vercel: Settings > Environment "
            "Variables) y vuelve a desplegar.",
        )

    uploaded = request.files.getlist("files")
    if not uploaded:
        raise AppError(400, "No se ha subido ningún archivo")
    if len(uploaded) > MAX_FILES:
        raise AppError(400, f"Máximo {MAX_FILES} archivos por lote")

    # Máster seleccionado (id o nombre; vacío = máster por defecto): decide a
    # qué formulario de valoración apunta el QR de la diapositiva de cierre.
    master_value = request.form.get("master", "").strip()
    master = resolve_master(master_value)
    if master is None:
        valid = ", ".join(m["name"] for m in list_masters())
        raise AppError(400, f"Máster no reconocido: {master_value!r}. Opciones: {valid}")
    qr_url = forms_url_for(master)

    master_path = _ensure_master_template()

    files_data = []
    skip_errors = []
    for f in uploaded:
        raw = f.read()
        if len(raw) > MAX_FILE_MB * 1024 * 1024:
            skip_errors.append(f"{f.filename}: supera los {MAX_FILE_MB}MB, se omite")
            continue
        files_data.append((f.filename, raw))

    results, errors = asyncio.run(_process_batch(files_data, api_key, master_path, qr_url))
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
