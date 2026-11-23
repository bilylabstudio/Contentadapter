"""
Extracción de contenido "en bruto" de la presentación/PDF antiguo que se
quiere adaptar. Produce, por cada diapositiva/página de origen, un dict:

{
  "index": int,
  "text": str,            # todo el texto de la diapositiva, en orden de lectura aproximado
  "text_blocks": [str,..] # bloques de texto separados (uno por forma/línea grande)
  "images": [{"bytes": bytes, "url": str|None}, ...]  # imágenes reales incrustadas,
                                                        # mayor a menor tamaño, con su
                                                        # hipervínculo si lo tenían (p.ej.
                                                        # miniatura de un vídeo de TikTok/YouTube)
}

Soporta .pptx (python-pptx) y .pdf (pdfplumber + pypdf).
"""
from __future__ import annotations

import hashlib
import io
from collections import defaultdict
from typing import BinaryIO

import pypdf
from PIL import Image
from pptx import Presentation


MIN_IMAGE_PIXELS = 120 * 120  # descarta iconos/logos minúsculos como "imagen de contenido"

# Si una imagen BYTE-IDÉNTICA aparece en 4 o más diapositivas/páginas distintas
# del mismo documento, es casi con toda seguridad un elemento de marca fijo
# (logo del máster, logo de la empresa del docente, pie de página, marca de
# agua...) y NO una foto de contenido específica de una diapositiva concreta.
# Se excluye por completo de las imágenes "candidatas a ilustrar" la
# diapositiva: así no ocupa el único hueco de imagen disponible desplazando a
# la foto real, ni se reutiliza como ilustración genérica repetida en
# diapositivas sin relación entre sí (quejas reales de un cliente: "sin el
# logo no se entiende el texto" en unas diapositivas, "hay muchas imágenes
# parecidas en todas las diapositivas" en otras -- ambas vienen de la misma
# causa). Si una diapositiva se queda sin ninguna imagen candidata tras este
# filtro, es preferible dejarla en su plantilla de texto plano a forzar un
# logo repetido como si fuera la foto de esa diapositiva.
_DECORATIVE_MIN_SLIDES = 4

# Tolerancia (como fracción del ancho de diapositiva) para considerar que dos
# imágenes comparten la misma posición/anchura horizontal -- es decir, que
# son "columnas" verticales del mismo elemento, no fotos distintas colocadas
# una al lado de la otra.
_STACK_X_TOL_FRAC = 0.03

# Tolerancia (como fracción de la altura del fragmento) para considerar que
# dos fragmentos están "pegados" verticalmente (sin hueco real entre ellos),
# permitiendo un pequeño solape o hueco como el que deja un recorte manual de
# una captura de pantalla larga.
_STACK_GAP_TOL_FRAC = 0.15


def _img_hash(blob: bytes) -> str:
    return hashlib.md5(blob).hexdigest()


def _shape_hyperlink(shp) -> str | None:
    """Hipervínculo (click para abrir) asociado a una forma, si lo tiene --
    p.ej. la miniatura de un vídeo de TikTok/YouTube enlazada al vídeo real.
    Nota: un enlace "ir a la diapositiva N" (salto interno) no es un
    hipervínculo (usa click_action.target_slide, no .hyperlink) y aquí
    siempre da None -- correcto, porque la numeración de diapositivas
    cambia por completo al adaptar y ese salto ya no tendría sentido."""
    try:
        url = shp.click_action.hyperlink.address
        return url or None
    except Exception:
        return None


def _shape_position(shp) -> tuple[int, int, int, int] | None:
    """Posición/tamaño (left, top, width, height) de una forma en EMU, si se
    puede leer. Se usa solo para detectar capturas de pantalla largas
    partidas en varios trozos apilados (ver `_merge_stacked_fragments`); si
    no está disponible (p.ej. forma dentro de un grupo con transform raro),
    esa imagen simplemente no participa en la fusión y sigue su camino
    normal -- no es un dato imprescindible para el resto del pipeline."""
    try:
        left, top, width, height = shp.left, shp.top, shp.width, shp.height
        if None in (left, top, width, height):
            return None
        return (int(left), int(top), int(width), int(height))
    except Exception:
        return None


def _chart_text(shp) -> str:
    """Resumen en texto de un gráfico nativo de PowerPoint (barras, líneas,
    tarta...): título + pares categoría/valor de cada serie. python-pptx no
    puede convertir un gráfico en imagen, así que sin esto el contenido de
    cualquier diapositiva con un gráfico (muy común: datos de encuestas,
    evolución de métricas...) desaparecía por completo -- ni como texto ni
    como imagen -- sin ningún aviso."""
    try:
        chart = shp.chart
    except Exception:
        return ""
    parts: list = []
    try:
        if chart.has_title and chart.chart_title.text_frame.text.strip():
            parts.append(chart.chart_title.text_frame.text.strip())
    except Exception:
        pass
    try:
        for plot in chart.plots:
            try:
                cats = [str(c) for c in plot.categories]
            except Exception:
                cats = []
            for series in plot.series:
                try:
                    vals = [str(v) if v is not None else "" for v in series.values]
                except Exception:
                    continue
                label = (series.name or "").strip()
                if cats and vals and len(cats) == len(vals):
                    pairs = ", ".join(f"{c}: {v}" for c, v in zip(cats, vals) if v)
                    if pairs:
                        parts.append(f"{label}: {pairs}" if label else pairs)
                elif vals:
                    joined = ", ".join(v for v in vals if v)
                    if joined:
                        parts.append(f"{label}: {joined}" if label else joined)
    except Exception:
        pass
    return "\n".join(p for p in parts if p)


def _walk_shapes(shapes, blocks: list, images: list, tables: list, depth: int = 0) -> None:
    """Recorre las formas de una diapositiva (incluidos grupos anidados) y
    recoge texto, imágenes (con su hipervínculo y posición si las tienen) y
    tablas reales."""
    for shp in shapes:
        try:
            if getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip():
                blocks.append(shp.text_frame.text.strip())
            if getattr(shp, "has_table", False) and shp.has_table:
                rows = []
                for row in shp.table.rows:
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    if any(cells):
                        rows.append(cells)
                if rows:
                    tables.append(rows)
                    for cells in rows:
                        blocks.append(" | ".join(c for c in cells if c))
            if shp.shape_type == 13:  # PICTURE
                try:
                    blob = shp.image.blob
                    if len(blob) > 8000:  # descarta imágenes triviales/decorativas muy pequeñas
                        images.append({
                            "bytes": blob,
                            "url": _shape_hyperlink(shp),
                            "pos": _shape_position(shp),
                        })
                except Exception:
                    pass
            if getattr(shp, "has_chart", False) and shp.has_chart:
                chart_txt = _chart_text(shp)
                if chart_txt:
                    blocks.append(chart_txt)
            if getattr(shp, "shape_type", None) == 6 and depth < 5:  # GROUP -- baja niveles
                _walk_shapes(shp.shapes, blocks, images, tables, depth + 1)
        except Exception:
            continue


def _merge_stacked_fragments(images: list[dict], ref_width: int) -> list[dict]:
    """Une en una sola imagen los grupos de imágenes que son, en realidad,
    trozos verticales de una misma captura de pantalla larga: misma posición
    horizontal (mismo `left` y `width`, dentro de tolerancia) y apiladas sin
    hueco real entre un trozo y el siguiente (`top` del trozo N+1 ≈ `top` +
    `height` del trozo N). Es un patrón muy habitual cuando alguien pega en
    la diapositiva una captura de pantalla de scroll largo (p.ej. resultados
    de Google) partida en 2-4 trozos porque no cabía entera en una sola
    captura.

    Sin esto, el resto del pipeline (que se queda solo con las 4 imágenes
    más grandes en bytes de la diapositiva, pensando que son fotos
    independientes) elegía arbitrariamente uno solo de los trozos como "la"
    imagen de la diapositiva, perdiendo el resto del contenido de la
    captura sin ningún aviso (ejemplo real: una captura de resultados de
    Google cortada por el lateral, con palabras partidas a mitad).

    Las imágenes SIN posición conocida (`pos` es None) nunca se agrupan --
    se dejan tal cual, igual que antes de este fix. Si algún fragmento del
    grupo no se puede decodificar con Pillow, ese grupo se deja intacto en
    vez de arriesgarse a perder contenido o producir un archivo roto."""
    positioned = [img for img in images if img.get("pos")]
    unpositioned = [img for img in images if not img.get("pos")]
    if len(positioned) < 2:
        return images

    x_tol = max(int(ref_width * _STACK_X_TOL_FRAC), 1) if ref_width else 10 ** 9

    n = len(positioned)
    used = [False] * n
    groups: list[list[int]] = []
    for i in range(n):
        if used[i]:
            continue
        group = [i]
        used[i] = True
        changed = True
        while changed:
            changed = False
            for j in range(n):
                if used[j]:
                    continue
                left_j, top_j, w_j, h_j = positioned[j]["pos"]
                aligned = any(
                    abs(positioned[k]["pos"][0] - left_j) <= x_tol
                    and abs(positioned[k]["pos"][2] - w_j) <= x_tol
                    for k in group
                )
                if not aligned:
                    continue
                tops = [positioned[k]["pos"][1] for k in group]
                heights = [positioned[k]["pos"][3] for k in group]
                bottoms = [t + h for t, h in zip(tops, heights)]
                gap_tol = int(max([h_j] + heights) * _STACK_GAP_TOL_FRAC)
                adjacent = (
                    any(abs(top_j - b) <= gap_tol for b in bottoms)
                    or any(abs((top_j + h_j) - t) <= gap_tol for t in tops)
                )
                if adjacent:
                    group.append(j)
                    used[j] = True
                    changed = True
        groups.append(group)

    merged: list[dict] = []
    for group in groups:
        if len(group) == 1:
            merged.append(positioned[group[0]])
            continue
        members = sorted((positioned[k] for k in group), key=lambda d: d["pos"][1])
        try:
            pil_imgs = [Image.open(io.BytesIO(m["bytes"])).convert("RGB") for m in members]
        except Exception:
            # No se pudo decodificar algún trozo: mejor dejarlos todos tal
            # cual (igual que antes de este fix) que arriesgar un resultado
            # corrupto o incompleto.
            merged.extend(members)
            continue
        target_w = max(im.width for im in pil_imgs)
        scaled = []
        for im in pil_imgs:
            if im.width != target_w:
                new_h = max(1, round(im.height * (target_w / im.width)))
                im = im.resize((target_w, new_h))
            scaled.append(im)
        total_h = sum(im.height for im in scaled)
        canvas = Image.new("RGB", (target_w, total_h), "white")
        y = 0
        for im in scaled:
            canvas.paste(im, (0, y))
            y += im.height
        buf = io.BytesIO()
        canvas.save(buf, format="JPEG", quality=92)
        url = next((m.get("url") for m in members if m.get("url")), None)
        left0, top0, w0, _ = members[0]["pos"]
        total_h_emu = sum(m["pos"][3] for m in members)
        merged.append({
            "bytes": buf.getvalue(),
            "url": url,
            "pos": (left0, top0, w0, total_h_emu),
            # Esta imagen es una composición sintética de varios trozos de
            # una captura de pantalla (texto denso: resultados de buscador,
            # reseñas...). El recorte "cover" de build.py (centrado, hasta un
            # 30% de la imagen) es aceptable para fotos pero en una imagen
            # así de texto-densa incluso un recorte pequeño corta palabras en
            # los bordes (p.ej. títulos de resultado que empiezan pegados al
            # borde izquierdo). Se marca para que build.py la trate siempre
            # en modo "contain" (nunca recortada), sea cual sea su proporción.
            "no_crop": True,
        })
    return merged + unpositioned


def _filter_decorative(raw_slides: list[dict]) -> None:
    """Modifica `raw_slides` in-place: calcula qué imágenes se repiten
    idénticas en >= _DECORATIVE_MIN_SLIDES diapositivas/páginas distintas del
    documento y las quita de `images` (se consideran marca/decoración fija,
    no contenido). Ordena las imágenes restantes de mayor a menor tamaño y
    recorta a 4 por diapositiva, igual que antes de este filtro."""
    hash_slides: dict[str, set] = defaultdict(set)
    for s in raw_slides:
        for img in s["images"]:
            hash_slides[_img_hash(img["bytes"])].add(s["index"])

    decorative = {h for h, idxs in hash_slides.items() if len(idxs) >= _DECORATIVE_MIN_SLIDES}

    for s in raw_slides:
        kept = [img for img in s["images"] if _img_hash(img["bytes"]) not in decorative]
        kept.sort(key=lambda d: len(d["bytes"]), reverse=True)
        s["images"] = kept[:4]


def extract_pptx(file_obj: BinaryIO) -> list[dict]:
    prs = Presentation(file_obj)
    slide_width = prs.slide_width or 0
    slides = []
    for i, slide in enumerate(prs.slides):
        blocks: list = []
        images: list = []
        tables: list = []
        _walk_shapes(slide.shapes, blocks, images, tables)
        images = _merge_stacked_fragments(images, slide_width)
        slides.append({
            "index": i,
            "text": "\n".join(blocks),
            "text_blocks": blocks,
            "images": images,
            "tables": tables,
        })
    _filter_decorative(slides)
    return slides


def extract_pdf(file_obj: BinaryIO) -> list[dict]:
    reader = pypdf.PdfReader(file_obj)
    pages = []
    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        blocks = [b.strip() for b in text.split("\n") if b.strip()]
        images = []
        try:
            for im in page.images:
                try:
                    w, h = im.image.size
                except Exception:
                    continue
                if w * h < MIN_IMAGE_PIXELS:
                    continue
                buf = io.BytesIO()
                im.image.convert("RGB").save(buf, format="JPEG", quality=90)
                images.append({"bytes": buf.getvalue(), "url": None})
        except Exception:
            pass
        pages.append({
            "index": i,
            "text": "\n".join(blocks),
            "text_blocks": blocks,
            "images": images,
        })
    _filter_decorative(pages)
    return pages


def extract(filename: str, file_obj: BinaryIO) -> list[dict]:
    name = filename.lower()
    if name.endswith(".pptx"):
        return extract_pptx(file_obj)
    if name.endswith(".pdf"):
        return extract_pdf(file_obj)
    raise ValueError(f"Formato no soportado: {filename} (solo .pptx y .pdf)")
