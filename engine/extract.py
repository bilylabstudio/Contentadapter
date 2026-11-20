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
    recoge texto, imágenes (con su hipervínculo si lo tienen) y tablas reales."""
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
                        images.append({"bytes": blob, "url": _shape_hyperlink(shp)})
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
    slides = []
    for i, slide in enumerate(prs.slides):
        blocks: list = []
        images: list = []
        tables: list = []
        _walk_shapes(slide.shapes, blocks, images, tables)
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
