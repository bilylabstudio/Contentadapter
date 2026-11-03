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
import re
import unicodedata
from collections import Counter
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
(¿es una lista? ¿una comparación? ¿un proceso? ¿una cita/dato? ¿una imagen?) -- EXCEPTO por la \
regla de imágenes de abajo, que tiene prioridad sobre esto para contenido de "idea única/lista".
- Si la diapositiva de origen es una portada/título general del documento, usa 'portada'. Si es \
un separador de bloque/tema, usa 'seccion'. Si es el cierre/agradecimiento, usa 'cierre'. Estas \
plantillas (junto con las de tarjetas, pasos, dos columnas, dato destacado, tabla y ejercicio) \
son intencionadamente planas y SIN foto -- así está diseñada la marca IMSED para esos casos, no \
es un fallo, así que no intentes forzar una imagen ahí.
- REGLA DE IMÁGENES (importante, es una queja recurrente del cliente: la presentación adaptada \
se queda "vacía y monótona" si no se reutilizan las fotos del documento original): si la \
diapositiva de origen tiene al menos una imagen real extraída Y el contenido encaja como "idea \
única" o "lista de puntos" (el tipo de contenido de 'contenido_bullets'/'contenido_simple'), \
usa SIEMPRE 'texto_imagen' o 'imagen_texto' en su lugar, poniendo esa imagen en 'image_index'. \
NO hace falta que la imagen sea un dato/gráfico "informativo" -- una foto de recurso o decorativa \
del documento original también cuenta y es preferible a una diapositiva plana sin imagen. Alterna \
'texto_imagen' (imagen a la derecha) e 'imagen_texto' (imagen a la izquierda) en diapositivas \
consecutivas con imagen para variar la composición. Usa 'imagen_completa' si la diapositiva de \
origen es casi solo una foto grande con poco texto. Esta regla NO aplica a las plantillas de \
tarjetas/pasos/dato destacado/etc. (ver punto anterior): si el contenido encaja mejor como \
tarjetas o pasos por su estructura, mantén esa plantilla aunque haya imagen disponible -- esas \
plantillas no tienen hueco de imagen por diseño.
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


# Plantillas de "idea única / lista de puntos" que SÍ tienen equivalente con
# imagen en la plantilla madre. Mapea cada una a su gemela con foto.
_IMAGE_UPGRADE = {"contenido_bullets", "contenido_simple"}

# Umbral de longitud (caracteres) del cuerpo de texto por debajo del cual es
# seguro forzar el cambio a una plantilla con imagen: esas plantillas usan
# una caja de texto a MEDIA anchura (el resto lo ocupa la foto), en vez de
# la anchura completa de 'contenido_bullets'/'contenido_simple'. El
# autofit de PowerPoint (MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, ver
# build.set_shape_text) encoge la fuente si hace falta, pero no es
# instantáneo en todos los visores y tiene un límite razonable de encogido
# antes de volverse ilegible. Si el texto es más largo que este umbral, NO
# forzamos el cambio de plantilla y dejamos el original a anchura completa
# -- evitar que el texto se salga de su caja o pise la imagen tiene
# prioridad sobre el reciclaje de fotos cuando ambos objetivos chocan.
_MAX_SAFE_BODY_CHARS_FOR_IMAGE_UPGRADE = 420


def _body_text_length(fields: list[dict]) -> int:
    """Suma la longitud del/de los campo(s) de cuerpo (pos '3', el mismo en
    contenido_bullets/contenido_simple/texto_imagen) de una lista de fields
    del LLM."""
    return sum(len(f.get("text") or "") for f in fields if f.get("pos") == "3")


def _enforce_image_reuse(raw: dict, slide: dict, use_left_side: bool) -> dict:
    """Refuerzo determinista de la regla "usa la foto del original" del
    prompt: el LLM no siempre la sigue de forma fiable (es una instrucción de
    texto, no una garantía), y el cliente ha pedido explícitamente más
    reciclaje de fotos porque el resultado se veía "vacío y monótono". Si
    Gemini elige una plantilla de texto plano para una diapositiva que SÍ
    tiene una imagen real extraída (y no marcó ya una imagen), la reescribimos
    aquí a la plantilla gemela con foto -- mismo contenido de texto, solo
    cambia el layout y se añade la imagen. No toca nada si el LLM ya siguió
    la regla, ni afecta a plantillas de tarjetas/pasos/etc. (esas son planas
    por diseño de marca, no se tocan)."""
    if raw.get("skip"):
        return raw
    if raw.get("template_id") not in _IMAGE_UPGRADE:
        return raw
    images = slide.get("images", [])
    if not images:
        return raw
    existing_idx = raw.get("image_index")
    if isinstance(existing_idx, int) and existing_idx >= 0:
        return raw  # el LLM ya usó una imagen, no hay nada que forzar

    fields = raw.get("fields", []) or []
    if _body_text_length(fields) > _MAX_SAFE_BODY_CHARS_FOR_IMAGE_UPGRADE:
        # el cuerpo es demasiado largo para la caja a media anchura de las
        # plantillas con imagen: mejor dejarlo en su plantilla original a
        # anchura completa que arriesgarse a que se salga de su caja o pise
        # la imagen
        return raw

    new_template_id = "imagen_texto" if use_left_side else "texto_imagen"
    if new_template_id == "imagen_texto":
        # imagen_texto usa la posición "4" para el cuerpo en vez de "3"
        new_fields = [
            {**f, "pos": "4"} if f.get("pos") == "3" else dict(f)
            for f in fields
        ]
    else:
        new_fields = [dict(f) for f in fields]

    return {
        **raw,
        "template_id": new_template_id,
        "fields": new_fields,
        "image_index": 0,
    }


def _resolve_spec(raw: dict, keep_empty: bool = False, kicker_default: str = "") -> Optional[dict]:
    """Convierte la respuesta del LLM (con template_id del catálogo) en un
    'render spec' que engine.build.build_deck entiende (con example_index,
    modo, delete_positions resueltos a partir de omit_groups, etc.)."""
    if raw.get("skip"):
        return None
    tpl = BY_ID.get(raw.get("template_id"))
    if tpl is None:
        return None

    if keep_empty:
        # modo fiel: los huecos sin contenido se dejan EN BLANCO (si no, la
        # plantilla conservaría su texto de ejemplo)
        fields = {f["pos"]: (f.get("text") or "") for f in raw.get("fields", [])}
    else:
        fields = {f["pos"]: f["text"] for f in raw.get("fields", []) if f.get("text")}

    delete_positions = []
    groups = tpl.get("groups")
    if groups:
        for gi in raw.get("omit_groups", []) or []:
            if 0 <= gi < len(groups):
                for pos in groups[gi]:
                    delete_positions.append(int(pos))
                    fields.pop(pos, None)

    if keep_empty:
        deleted = {str(p) for p in delete_positions}
        for pos, desc in tpl["editable"].items():
            if pos in fields and fields[pos]:
                continue
            if pos in deleted:
                continue
            if str(desc).lower().startswith("kicker") and kicker_default:
                fields[pos] = kicker_default
            else:
                fields.setdefault(pos, "")

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


async def _call_gemini(client: httpx.AsyncClient, api_key: str, system_prompt: str, user_prompt: str, temperature: float = 0.4) -> dict:
    payload = {
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "systemInstruction": {"role": "system", "parts": [{"text": system_prompt}]},
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": RESPONSE_SCHEMA,
            "temperature": temperature,
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
                    results[i] = {"_error": _clean_err(e), "_slide_index": slide["index"]}
                    return
                raw = _enforce_image_reuse(raw, slide, use_left_side=(i % 2 == 1))
                spec = _resolve_spec(raw)
                if spec is not None:
                    _attach_image(spec, slide)
                results[i] = spec

        await asyncio.gather(*(worker(i, s) for i, s in enumerate(slides)))
    return results


# ---------------------------------------------------------------------------
# Utilidades comunes
# ---------------------------------------------------------------------------

def _clean_err(e) -> str:
    """Mensaje de error de Gemini SIN la API key (la URL de la petición la
    lleva en el query string) y con un texto claro para el caso de falta de
    saldo."""
    msg = str(e)
    msg = re.sub(r"key=[^&'\s\"]+", "key=***", msg)
    if "402" in msg and "Payment Required" in msg:
        return ("Gemini rechazó la petición por falta de saldo o facturación (402). "
                "Avisa al administrador para recargar la cuenta de Gemini.")
    if "429" in msg and "Too Many Requests" in msg:
        return "Gemini ha alcanzado su límite de peticiones (429). Vuelve a intentarlo en unos minutos."
    return msg[:300]


def _attach_image(spec: dict, slide: dict) -> None:
    """Resuelve la imagen elegida (índice) a bytes reales de la diapositiva de origen."""
    img_idx = spec.pop("_image_index", None)
    if img_idx is not None:
        images = slide.get("images", [])
        if 0 <= img_idx < len(images):
            field = spec.pop("_image_field", None) or spec.pop("_image_slot", None)
            if field is not None:
                spec["images"][str(field)] = images[img_idx]
    spec.pop("_image_field", None)
    spec.pop("_image_slot", None)


# ---------------------------------------------------------------------------
# MODO FIEL: traspaso del contenido a la plantilla IMSED sin reescribirlo
# ---------------------------------------------------------------------------

# Plantillas permitidas en modo fiel: las que admiten texto libre o que
# reproducen una estructura de origen sin tener que resumir. Se excluyen las
# de tarjetas/pasos/dato destacado/ejercicio, que obligan a condensar.
FIEL_TEMPLATES = {
    "portada", "seccion", "contenido_bullets", "contenido_simple",
    "texto_imagen", "imagen_texto", "imagen_completa", "tabla",
    "dos_columnas", "cierre",
}
_BODY_POS = {"contenido_bullets": "3", "contenido_simple": "3", "texto_imagen": "3", "imagen_texto": "4"}
_IMAGE_TEMPLATES = {"texto_imagen", "imagen_texto"}
MAX_BODY_FULL = 650        # caracteres de cuerpo por diapositiva a anchura completa
MAX_BODY_WITH_IMAGE = 420  # idem con imagen (caja a media anchura)
COVERAGE_MIN = 0.98        # % mínimo de palabras de origen presentes en el resultado
EXTRA_MAX = 0.08           # % máximo de palabras del resultado que no estaban en el origen

FIEL_SYSTEM_PROMPT = """Eres el motor de una herramienta que TRASPASA presentaciones antiguas de IMSED \
(másteres/formación) a una plantilla madre nueva SIN reescribir su contenido. El profesor quiere que \
se respete SU contenido: solo cambia el diseño.

Reglas (obligatorias):
- COPIA el texto de origen LITERALMENTE: mismas palabras y mismo orden. No resumas, no reformules, no \
acortes, no corrijas, no traduzcas, no añadas frases ni datos nuevos. No pierdas ninguna frase ni cifra.
- Solo puedes: (a) repartir los bloques de texto de origen entre los huecos de la plantilla elegida \
(título, subtítulo, cuerpo...), (b) separar con \\n las líneas o puntos que ya estaban separados en el \
origen, (c) quitar el símbolo de viñeta (•, -, ·...) del inicio de una línea.
- Una diapositiva de origen = UNA diapositiva de destino. No marques skip salvo que la diapositiva no \
tenga NINGÚN texto ni imagen.
- El título es el bloque corto del principio (si lo hay); el resto va en el cuerpo, en el orden original. \
Deja en blanco ("") los huecos para los que el origen no tiene texto (subtítulo, etc.). No inventes \
subtítulos ni kickers.
- Elige la plantilla por la estructura del contenido: lista de puntos -> 'contenido_bullets'; párrafo o \
texto corrido -> 'contenido_simple'; comparativa de dos columnas -> 'dos_columnas'; tabla real -> 'tabla' \
(todas las filas y columnas, sin recortar); portada -> 'portada'; separador de tema -> 'seccion'; \
despedida -> 'cierre'; foto grande casi sin texto -> 'imagen_completa'. Ante la duda usa \
'contenido_bullets' o 'contenido_simple', que admiten cualquier texto.
- Si la diapositiva de origen tiene una imagen real extraída y el contenido es una lista o párrafo, usa \
'texto_imagen' o 'imagen_texto' con image_index 0 (alterna lados entre diapositivas consecutivas).
- Devuelve SIEMPRE 'fields' como lista de {pos, text} usando EXACTAMENTE las posiciones del catálogo para \
la plantilla elegida.

CATÁLOGO DE PLANTILLAS DISPONIBLES:
{catalog}
"""

_BULLET_RE = re.compile(r"^[\s•·▪●◦■□○\-–—*➢➤►✓✔]+")


def _words(text: str) -> list[str]:
    s = unicodedata.normalize("NFKD", (text or "").lower())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.findall(r"[a-z0-9]+", s)


def _raw_texts(raw: dict) -> tuple[str, str]:
    """(todo el texto del resultado, texto sin kicker) de un mapeo del LLM."""
    all_t, nok_t = [], []
    for f in raw.get("fields", []) or []:
        t = f.get("text") or ""
        all_t.append(t)
        if str(f.get("pos")) != "0":
            nok_t.append(t)
    for row in raw.get("table_rows", []) or []:
        for cell in row:
            all_t.append(str(cell))
            nok_t.append(str(cell))
    return " ".join(all_t), " ".join(nok_t)


def _fiel_check(raw: dict, slide: dict) -> tuple[bool, float]:
    """¿Conserva el mapeo del LLM el contenido de origen? Devuelve (ok, cobertura)."""
    src = Counter(_words(slide.get("text", "")))
    n_src = sum(src.values())
    if n_src == 0:
        return True, 1.0
    if raw.get("skip"):
        return (n_src <= 2), 0.0
    all_t, nok_t = _raw_texts(raw)
    out_all = Counter(_words(all_t))
    out_nok = Counter(_words(nok_t))
    coverage = sum((src & out_all).values()) / n_src
    n_out = sum(out_nok.values())
    extra = (sum((out_nok - src).values()) / n_out) if n_out else 0.0
    return (coverage >= COVERAGE_MIN and extra <= EXTRA_MAX), coverage


def _literal_raw(slide: dict) -> Optional[dict]:
    """Traspaso 100 % literal sin LLM: primera línea corta = título, el resto = cuerpo."""
    lines = []
    for b in slide.get("text_blocks") or slide.get("text", "").split("\n"):
        for l in str(b).split("\n"):
            l = _BULLET_RE.sub("", l).strip()
            if l:
                lines.append(l)
    images = slide.get("images", [])
    if not lines:
        return None
    title = ""
    body = lines
    if len(lines) >= 2 and len(lines[0]) <= 100:
        title, body = lines[0], lines[1:]
    body_txt = "\n".join(body)
    if images and len(body_txt) <= MAX_BODY_WITH_IMAGE:
        return {"skip": False, "template_id": "texto_imagen", "image_index": 0,
                "fields": [{"pos": "1", "text": title}, {"pos": "3", "text": body_txt}]}
    tid = "contenido_bullets" if len(body) > 1 else "contenido_simple"
    return {"skip": False, "template_id": tid, "image_index": -1,
            "fields": [{"pos": "1", "text": title}, {"pos": "3", "text": body_txt}]}


def _split_text(text: str, limit: int) -> list[str]:
    """Parte un texto en trozos <= limit sin cortar líneas (y, si una línea es
    más larga que el límite, sin cortar frases ni palabras)."""
    pieces = []
    for line in text.split("\n"):
        if len(line) <= limit:
            pieces.append(line)
            continue
        cur = ""
        for sent in re.split(r"(?<=[.!?;:])\s+", line):
            while len(sent) > limit:
                cut = sent.rfind(" ", 0, limit)
                cut = cut if cut > 0 else limit
                if cur:
                    pieces.append(cur); cur = ""
                pieces.append(sent[:cut].strip())
                sent = sent[cut:].strip()
            if cur and len(cur) + 1 + len(sent) > limit:
                pieces.append(cur); cur = sent
            else:
                cur = (cur + " " + sent).strip()
        if cur:
            pieces.append(cur)
    chunks, cur = [], ""
    for p in pieces:
        if cur and len(cur) + 1 + len(p) > limit:
            chunks.append(cur); cur = p
        else:
            cur = (cur + "\n" + p) if cur else p
    if cur:
        chunks.append(cur)
    return chunks or [text]


def _split_overflow(raw: dict) -> list[dict]:
    """Si el cuerpo es demasiado largo para su caja, lo reparte en diapositivas
    de continuación (mismo título + '(cont.)') en vez de recortar o encoger."""
    tid = raw.get("template_id")
    pos = _BODY_POS.get(tid)
    if pos is None:
        return [raw]
    fields = {str(f.get("pos")): (f.get("text") or "") for f in raw.get("fields", [])}
    body = fields.get(pos, "")
    limit = MAX_BODY_WITH_IMAGE if tid in _IMAGE_TEMPLATES else MAX_BODY_FULL
    if len(body) <= limit:
        return [raw]
    chunks = _split_text(body, limit)
    if len(chunks) == 1:
        return [raw]
    out = []
    for k, chunk in enumerate(chunks):
        if k == 0:
            f = dict(fields); f[pos] = chunk
            out.append({**raw, "fields": [{"pos": p, "text": t} for p, t in f.items()]})
        else:
            title = fields.get("1", "")
            out.append({
                "skip": False,
                "template_id": "contenido_bullets" if "\n" in chunk else "contenido_simple",
                "image_index": -1,
                "fields": [
                    {"pos": "0", "text": fields.get("0", "")},
                    {"pos": "1", "text": (title + " (cont.)") if title else ""},
                    {"pos": "2", "text": ""},
                    {"pos": "3", "text": chunk},
                ],
            })
    return out


def _short_master_name(name: str) -> str:
    name = (name or "").split(":")[0].strip()
    return name[:40]


def _build_fiel_prompt(slide: dict, deck_context: str) -> str:
    n_images = len(slide.get("images", []))
    img_line = (f"Imágenes reales disponibles de esta diapositiva: {n_images} (usa image_index 0 si procede)."
                if n_images else "No hay imágenes reales extraídas de esta diapositiva.")
    blocks = slide.get("text_blocks") or [slide.get("text", "")]
    numbered = "\n".join(f"[{i + 1}] {b}" for i, b in enumerate(blocks)) or "(sin texto)"
    return (
        f"Documento: {deck_context}\n\n"
        f"Bloques de texto de ESTA diapositiva de origen, en orden de lectura (cópialos literalmente):\n"
        f"---\n{numbered}\n---\n\n{img_line}\n\nDevuelve el JSON de mapeo para esta diapositiva."
    )


async def map_deck_fiel(slides: list[dict], filename: str, api_key: str, master_name: str = "") -> tuple[list[dict], dict]:
    """Modo fiel. Devuelve (specs de render en orden, informe). Una diapositiva
    de origen produce 1 diapositiva (o más si el texto no cabe y se reparte
    en continuaciones). Si el LLM altera el contenido o falla, esa diapositiva
    se traspasa de forma literal sin LLM. Si TODAS las llamadas a Gemini
    fallan, devuelve entradas '_error' (como el modo adaptado) para que la
    app avise en vez de entregar un resultado sin pasar por el modelo."""
    deck_context = _guess_deck_context(slides, filename)
    kicker_default = _short_master_name(master_name)
    system_prompt = FIEL_SYSTEM_PROMPT.replace("{catalog}", catalog_prompt_text(FIEL_TEMPLATES))
    sem = asyncio.Semaphore(MAX_CONCURRENCY)
    per_slide: list[list[dict]] = [[] for _ in slides]
    stats = {"gemini_ok": 0, "gemini_error": 0, "literal": 0, "divididas": 0, "omitidas": 0}
    first_error = {"msg": ""}

    async with httpx.AsyncClient() as client:
        async def worker(i: int, slide: dict):
            async with sem:
                raw = None
                try:
                    raw = await _call_gemini(client, api_key, system_prompt, _build_fiel_prompt(slide, deck_context), temperature=0.1)
                    stats["gemini_ok"] += 1
                except Exception as e:
                    stats["gemini_error"] += 1
                    first_error["msg"] = first_error["msg"] or _clean_err(e)
                if raw is not None and raw.get("template_id") not in FIEL_TEMPLATES and not raw.get("skip"):
                    raw = None  # plantilla no permitida en modo fiel
                ok = False
                if raw is not None:
                    ok, _cov = _fiel_check(raw, slide)
                if raw is not None and ok:
                    if raw.get("skip"):
                        stats["omitidas"] += 1
                        return
                    raw = _enforce_image_reuse(raw, slide, use_left_side=(i % 2 == 1))
                else:
                    raw = _literal_raw(slide)
                    if raw is None:
                        stats["omitidas"] += 1
                        return
                    stats["literal"] += 1
                parts = _split_overflow(raw)
                if len(parts) > 1:
                    stats["divididas"] += 1
                for part in parts:
                    spec = _resolve_spec(part, keep_empty=True, kicker_default=kicker_default)
                    if spec is not None:
                        _attach_image(spec, slide)
                        per_slide[i].append(spec)

        await asyncio.gather(*(worker(i, s) for i, s in enumerate(slides)))

    if slides and stats["gemini_error"] == len(slides):
        return [{"_error": first_error["msg"] or "Gemini no respondió", "_slide_index": s["index"]} for s in slides], stats
    specs = [sp for group in per_slide for sp in group]
    stats["diapositivas_origen"] = len(slides)
    stats["diapositivas_resultado"] = len(specs)
    if first_error["msg"]:
        stats["aviso"] = first_error["msg"]
    return specs, stats
