"""
Paso de "inteligencia": para cada diapositiva/página extraída del documento
antiguo, decide con Gemini qué plantilla de la plantilla madre IMSED encaja
mejor y qué texto va en cada hueco, respetando el estilo y los límites de
la plantilla original (una idea por diapositiva, frases cortas, etc.).

Stateless: no se guarda nada en disco ni en BBDD. Solo se llama a la API de
Gemini con el contenido de cada diapositiva y se recibe de vuelta un JSON.
"""
from __future__ import annotations

import asyncio
import json
import os
from typing import Optional

import httpx

from .catalog import BY_ID, catalog_prompt_text

# "gemini-2.0-flash" fue retirado por Google (devolvía 404 Not Found en
# generateContent). Usamos el alias "gemini-flash-latest", que Google
# redirige automáticamente al último modelo Flash estable -- así no
# volvemos a romper la app cada vez que retiren una versión concreta.
# Se puede fijar una versión exacta vía la variable de entorno GEMINI_MODEL
# si en el futuro hiciera falta.
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-latest")
GEMINI_URL = (
    f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
)
MAX_CONCURRENCY = int(os.environ.get("MAPPER_CONCURRENCY", "6"))

RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "skip": {"type": "BOOLEAN", "description": "true si esta diapositiva no aporta contenido útil (portada duplicada, separador vacío, índice repetido...) y se puede omitir"},
        "template_id": {"type": "STRING", "description": "uno de los id del catálogo de plantillas"},
        "fields": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "pos": {"type": "STRING"},
                    "text": {"type": "STRING"},
                },
                "required": ["pos", "text"],
            },
        },
        "image_index": {"type": "INTEGER", "description": "índice (0-based) de la imagen de origen a usar, o -1 si no hay imagen o no aplica"},
        "omit_groups": {"type": "ARRAY", "items": {"type": "INTEGER"}, "description": "para plantillas de tarjetas/pasos repetidos: índices (0-based) de las unidades que sobran y hay que omitir"},
        "table_rows": {"type": "ARRAY", "items": {"type": "ARRAY", "items": {"type": "STRING"}}, "description": "solo para la plantilla 'tabla': filas de la tabla, primera fila = cabecera"},
    },
    "required": ["skip", "template_id", "fields"],
}

SYSTEM_PROMPT = """Eres el motor de una herramienta que adapta presentaciones antiguas de IMSED \
(másteres/formación) al diseño de una plantilla madre nueva, conservando el contenido pero \
reescribiéndolo para que encaje en el hueco de texto de la plantilla elegida.

Reglas de la plantilla madre (léelas y respétalas):
- Una idea por diapositiva. Si el contenido de origen mezcla varias ideas grandes, quédate con \
la idea principal de ESTA diapositiva de origen (habrá una llamada por cada una).
- Frases cortas y directas, en español, tono cercano pero profesional (es formación para adultos).
- Nunca inventes datos, cifras o nombres que no estén en el texto de origen. Si algo no está \
claro, generaliza en vez de inventar.
- Respeta los límites de longitud indicados en cada plantilla (son aproximados, pero no los \
multipliques por más de 1.5x).
- Elige la plantilla (template_id) que estructuralmente mejor encaje con el contenido de origen, \
no la que más se parezca visualmente al original: lo importante es el TIPO de contenido \
(¿es una lista? ¿una comparación? ¿un proceso? ¿una cita/dato? ¿una imagen?).
- Si la diapositiva de origen es una portada/título general del documento, usa 'portada'. Si es \
un separador de bloque/tema, usa 'seccion'. Si es el cierre/agradecimiento, usa 'cierre'.
- Si hay una imagen de origen realmente informativa (una captura, foto, gráfico) y existe una \
plantilla con hueco de imagen, considera usarla -- pero el texto sigue siendo lo prioritario.
- Devuelve SIEMPRE 'fields' como lista de {pos, text} usando EXACTAMENTE las posiciones (pos) \
que se indican en el catálogo para la plantilla elegida. No inventes posiciones nuevas.
- Si la plantilla es de tarjetas/pasos repetidos y tienes menos elementos de los que caben, \
indica en 'omit_groups' los índices (0-based, en el orden en que aparecen en el catálogo) de \
las unidades que sobran -- nunca dejes una tarjeta con el texto de ejemplo original.
- Si de verdad la diapositiva no aporta nada (portada duplicada, página en blanco, separador \
sin contenido real), marca skip=true y devuelve fields vacío.

CATÁLOGO DE PLANTILLAS DISPONIBLES:
{catalog}
"""


def _build_user_prompt(slide: dict, deck_context: str) -> str:
    n_images = len(slide.get("images", []))
    img_line = (
        f"Imágenes disponibles de esta diapositiva de origen: {n_images} "
        f"(referéncialas por su índice 0-based en 'image_index'; -1 si ninguna aplica)."
        if n_images else "No hay imágenes reales extraídas de esta diapositiva."
    )
    text = slide.get("text", "").strip() or "(sin texto extraído)"
    return (
        f"Contexto general del documento de origen: {deck_context}\n\n"
        f"Contenido de ESTA diapositiva/página de origen (texto tal cual se extrajo, "
        f"el orden de líneas puede no ser perfecto):\n---\n{text}\n---\n\n"
        f"{img_line}\n\n"
        f"Devuelve el JSON de mapeo para esta diapositiva."
    )


def _guess_deck_context(slides: list[dict], filename: str) -> str:
    for s in slides[:3]:
        for line in s.get("text_blocks", []):
            if 15 <= len(line) <= 90 and line.strip():
                return line.strip()
    return filename.rsplit(".", 1)[0]


def _resolve_spec(raw: dict) -> Optional[dict]:
    """Convierte la respuesta del LLM (con template_id del catálogo) en un
    'render spec' que engine.build.build_deck entiende (con example_index,
    modo, delete_positions resueltos a partir de omit_groups, etc.)."""
    if raw.get("skip"):
        return None
    tpl = BY_ID.get(raw.get("template_id"))
    if tpl is None:
        return None

    fields = {f["pos"]: f["text"] for f in raw.get("fields", []) if f.get("text")}

    delete_positions = []
    groups = tpl.get("groups")
    if groups:
        for gi in raw.get("omit_groups", []) or []:
            if 0 <= gi < len(groups):
                for pos in groups[gi]:
                    delete_positions.append(int(pos))
                    fields.pop(pos, None)

    if tpl.get("placeholder_layout"):
        spec = {
            "mode": "placeholder",
            "layout": tpl["layout"],
            "fields": fields,
            "images": {},
        }
        image_field = tpl.get("image_field")
        image_index = raw.get("image_index", -1)
        if image_field is not None and isinstance(image_index, int) and image_index >= 0:
            spec["_image_index"] = image_index
            spec["_image_field"] = image_field
        return spec

    spec = {
        "mode": "duplicate",
        "example_index": tpl["example_index"],
        "shape_text": fields,
        "images": {},
        "delete_positions": delete_positions,
    }

    image_slot = tpl.get("image_slot")
    image_index = raw.get("image_index", -1)
    if image_slot is not None and isinstance(image_index, int) and image_index >= 0:
        spec["_image_index"] = image_index
        spec["_image_slot"] = image_slot
        for extra in tpl.get("delete_if_no_image", []):
            if int(extra) in delete_positions:
                delete_positions.remove(int(extra))
    else:
        for extra in tpl.get("delete_if_no_image", []):
            if int(extra) not in delete_positions:
                delete_positions.append(int(extra))

    table_slot = tpl.get("table_slot")
    if table_slot and raw.get("table_rows"):
        spec["table"] = {"pos": int(table_slot), "rows": raw["table_rows"]}

    return spec


async def _call_gemini(client: httpx.AsyncClient, api_key: str, system_prompt: str, user_prompt: str) -> dict:
    payload = {
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "systemInstruction": {"role": "system", "parts": [{"text": system_prompt}]},
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": RESPONSE_SCHEMA,
            "temperature": 0.4,
        },
    }
    resp = await client.post(
        GEMINI_URL,
        params={"key": api_key},
        json=payload,
        timeout=60.0,
    )
    resp.raise_for_status()
    data = resp.json()
    text = data["candidates"][0]["content"]["parts"][0]["text"]
    return json.loads(text)


async def map_deck(slides: list[dict], filename: str, api_key: str) -> list[Optional[dict]]:
    """Devuelve una lista (mismo orden/longitud que `slides`) de render-specs
    (o None si esa diapositiva se omite), con las imágenes YA resueltas a
    bytes reales (tomadas de slide['images'][image_index])."""
    deck_context = _guess_deck_context(slides, filename)
    system_prompt = SYSTEM_PROMPT.replace("{catalog}", catalog_prompt_text())
    sem = asyncio.Semaphore(MAX_CONCURRENCY)
    results: list[Optional[dict]] = [None] * len(slides)

    async with httpx.AsyncClient() as client:
        async def worker(i: int, slide: dict):
            async with sem:
                user_prompt = _build_user_prompt(slide, deck_context)
                try:
                    raw = await _call_gemini(client, api_key, system_prompt, user_prompt)
                except Exception as e:
                    results[i] = {"_error": str(e), "_slide_index": slide["index"]}
                    return
                spec = _resolve_spec(raw)
                if spec is not None:
                    img_idx = spec.pop("_image_index", None)
                    if img_idx is not None:
                        images = slide.get("images", [])
                        if 0 <= img_idx < len(images):
                            field = spec.pop("_image_field", None) or spec.pop("_image_slot", None)
                            if field is not None:
                                spec["images"][str(field)] = images[img_idx]
                    spec.pop("_image_field", None)
                    spec.pop("_image_slot", None)
                results[i] = spec

        await asyncio.gather(*(worker(i, s) for i, s in enumerate(slides)))
    return results
