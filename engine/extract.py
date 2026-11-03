"""
Extracción de contenido "en bruto" de la presentación/PDF antiguo que se
quiere adaptar. Produce, por cada diapositiva/página de origen, un dict:

{
  "index": int,
  "text": str,            # todo el texto de la diapositiva, en orden de lectura aproximado
  "text_blocks": [str,..] # bloques de texto separados (uno por forma/línea grande)
  "images": [bytes, ...]  # imágenes reales incrustadas, como bytes (mayor a menor tamaño)
}

Soporta .pptx (python-pptx) y .pdf (pdfplumber + pypdf).
"""
from __future__ import annotations

import io
from typing import BinaryIO

import pypdf
from pptx import Presentation


MIN_IMAGE_PIXELS = 120 * 120  # descarta iconos/logos minúsculos como "imagen de contenido"


def _walk_shapes(shapes, blocks: list, images: list, tables: list, depth: int = 0) -> None:
    """Recorre las formas de una diapositiva (incluidos grupos anidados) y
    recoge texto, imágenes y tablas reales."""
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
                        images.append(blob)
                except Exception:
                    pass
            if getattr(shp, "shape_type", None) == 6 and depth < 5:  # GROUP -- baja niveles
                _walk_shapes(shp.shapes, blocks, images, tables, depth + 1)
        except Exception:
            continue


def extract_pptx(file_obj: BinaryIO) -> list[dict]:
    prs = Presentation(file_obj)
    slides = []
    for i, slide in enumerate(prs.slides):
        blocks: list = []
        images: list = []
        tables: list = []
        _walk_shapes(slide.shapes, blocks, images, tables)
        images.sort(key=len, reverse=True)
        slides.append({
            "index": i,
            "text": "\n".join(blocks),
            "text_blocks": blocks,
            "images": images[:4],
            "tables": tables,
        })
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
                images.append(buf.getvalue())
        except Exception:
            pass
        images.sort(key=len, reverse=True)
        pages.append({
            "index": i,
            "text": "\n".join(blocks),
            "text_blocks": blocks,
            "images": images[:4],
        })
    return pages


def extract(filename: str, file_obj: BinaryIO) -> list[dict]:
    name = filename.lower()
    if name.endswith(".pptx"):
        return extract_pptx(file_obj)
    if name.endswith(".pdf"):
        return extract_pdf(file_obj)
    raise ValueError(f"Formato no soportado: {filename} (solo .pptx y .pdf)")
