"""
Motor de renderizado: construye una presentación adaptada a partir de la
plantilla madre IMSED + una lista de "slide specs" (la salida del paso de
mapeo/clasificación, normalmente generado por un LLM).

Dos modos de construcción de diapositiva, según el layout de destino:

1. PLACEHOLDER-based (PORTADA, CONTENIDO, SECCION, CIERRE, DATO DESTACADO,
   TABLA*, PASOS*, DOS COLUMNAS*, DOS COLUMNAS*):
   se añade una diapositiva nueva a partir del layout y se rellenan sus
   placeholders de texto directamente. Rápido y siempre válido, pero no
   reproduce elementos "hechos a mano" (tarjetas, iconos, líneas...).

2. DUPLICATE-EXAMPLE (TARJETAS, PASOS, DOS COLUMNAS, EJERCICIO...):
   se duplica la diapositiva de EJEMPLO más parecida (ya tiene las formas,
   colores y posiciones correctas) y solo se sustituye el texto de cada
   forma editable, indexado por su posición en el orden de formas de esa
   diapositiva. Esto es lo que de verdad "hereda el diseño" tal y como pide
   IMSED.

Para las diapositivas con imagen (TEXTO + IMAGEN, IMAGEN + TEXTO,
IMAGEN COMPLETA) se borra el placeholder de imagen y se inserta la imagen
real en su mismo bounding box (ajustada por 'cover': se recorta para llenar
el hueco sin deformarla).
"""
from __future__ import annotations

import copy
import io
from dataclasses import dataclass, field
from typing import Optional

from PIL import Image
from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu


# ---------------------------------------------------------------------------
# Utilidades de bajo nivel
# ---------------------------------------------------------------------------

def delete_all_slides(prs: Presentation) -> None:
    """Vacía la presentación de diapositivas, dejando masters/layouts intactos."""
    xml_slides = prs.slides._sldIdLst
    slide_ids = list(xml_slides)
    for sld in slide_ids:
        rId = sld.get(qn("r:id"))
        prs.part.drop_rel(rId)
        xml_slides.remove(sld)


def _relink_images(source_slide, dest_slide) -> None:
    """Tras copiar formas por XML, re-crea las relaciones r:embed / r:link
    de imágenes para que apunten a partes válidas en la diapositiva destino."""
    src_part = source_slide.part
    dst_part = dest_slide.part
    for el in dest_slide.shapes._spTree.iter():
        for attr in ("embed", "link"):
            rId = el.get(qn(f"r:{attr}"))
            if not rId:
                continue
            try:
                rel = src_part.rels[rId]
            except KeyError:
                continue
            if rel.is_external:
                new_rId = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_rId = dst_part.relate_to(rel.target_part, rel.reltype)
            el.set(qn(f"r:{attr}"), new_rId)


def duplicate_slide(prs: Presentation, index: int):
    """Duplica prs.slides[index] DENTRO de la misma presentación (misma plantilla),
    devolviendo la nueva diapositiva. Preserva formas, estilos e imágenes."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)

    # quita los placeholders que add_slide crea automáticamente heredados del layout
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    # copia cada forma de la diapositiva de ejemplo, en el mismo orden
    for shp in source.shapes:
        new_el = copy.deepcopy(shp._element)
        dest.shapes._spTree.append(new_el)

    _relink_images(source, dest)
    return dest


def list_editable_shapes(slide) -> list[dict]:
    """Devuelve, en orden de documento, las formas de una diapositiva que
    contienen texto editable (placeholders y cajas de texto), con su índice
    posicional -- este índice es el identificador estable que usa el paso de
    mapeo para indicar qué forma reemplazar."""
    out = []
    for i, shp in enumerate(slide.shapes):
        if shp.has_text_frame and shp.text_frame.text.strip():
            out.append({"pos": i, "text": shp.text_frame.text, "name": shp.name})
    return out


def set_shape_text(shape, text: str) -> None:
    """Sustituye el texto de una forma preservando el formato COMPLETO de
    cada párrafo original: no solo fuente/tamaño/color, sino también su
    <a:pPr> (viñetas, sangría, interlineado...). Si el texto nuevo tiene más
    líneas que párrafos originales, cada línea de más clona el <a:pPr> del
    párrafo original más cercano (el último), para que las viñetas de una
    lista se repitan correctamente en vez de perderse."""
    tf = shape.text_frame
    txBody = tf._txBody
    orig_ps = txBody.findall(qn("a:p"))
    new_lines = text.split("\n")

    if not orig_ps:
        tf.text = text
        return

    # plantillas: copia de cada <a:p> original (con su pPr intacto) y del
    # rPr del primer run de cada uno (para tipografía/tamaño/color/negrita)
    p_templates = [copy.deepcopy(p) for p in orig_ps]
    rpr_templates = []
    for p in p_templates:
        r = p.find(qn("a:r"))
        rPr = r.find(qn("a:rPr")) if r is not None else None
        rpr_templates.append(copy.deepcopy(rPr) if rPr is not None else None)

    for p in orig_ps:
        txBody.remove(p)

    for i, line in enumerate(new_lines):
        t_idx = min(i, len(p_templates) - 1)
        new_p = copy.deepcopy(p_templates[t_idx])
        for r in new_p.findall(qn("a:r")):
            new_p.remove(r)
        for br in new_p.findall(qn("a:br")):
            new_p.remove(br)

        r_el = new_p.makeelement(qn("a:r"), {})
        rPr = rpr_templates[t_idx]
        if rPr is not None:
            r_el.append(copy.deepcopy(rPr))
        t_el = r_el.makeelement(qn("a:t"), {})
        t_el.text = line
        r_el.append(t_el)

        pPr = new_p.find(qn("a:pPr"))
        if pPr is not None:
            pPr.addnext(r_el)
        else:
            new_p.insert(0, r_el)

        txBody.append(new_p)

    # evita que texto más largo que el original desborde la forma y pise a
    # las formas vecinas (título largo tapando el subtítulo, etc.): que
    # PowerPoint/LibreOffice reduzca el tamaño de fuente automáticamente.
    try:
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass


def remove_shape_at(slide, pos: int) -> None:
    shapes = list(slide.shapes)
    if 0 <= pos < len(shapes):
        shp = shapes[pos]
        shp._element.getparent().remove(shp._element)


def insert_image_cover(slide, box, image_bytes: bytes) -> None:
    """Inserta una imagen ocupando exactamente `box` (left, top, width, height
    en EMU), recortándola (cover-fit) para llenar el hueco sin deformarla."""
    left, top, width, height = box
    im = Image.open(io.BytesIO(image_bytes))
    im = im.convert("RGB")
    src_w, src_h = im.size
    target_ratio = width / height
    src_ratio = src_w / src_h
    if src_ratio > target_ratio:
        # imagen más ancha de lo necesario -> recorta lados
        new_w = int(src_h * target_ratio)
        x0 = (src_w - new_w) // 2
        im = im.crop((x0, 0, x0 + new_w, src_h))
    else:
        new_h = int(src_w / target_ratio)
        y0 = (src_h - new_h) // 2
        im = im.crop((0, y0, src_w, y0 + new_h))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, width, height)


def get_placeholder(slide, idx: int):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def get_placeholder_box(layout, idx: int):
    for ph in layout.placeholders:
        if ph.placeholder_format.idx == idx:
            return (ph.left, ph.top, ph.width, ph.height)
    return None


# ---------------------------------------------------------------------------
# Construcción de diapositivas a partir de una "slide spec"
# ---------------------------------------------------------------------------

def _find_layout(prs: Presentation, name: str):
    for layout in prs.slide_masters[0].slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"Layout no encontrado en la plantilla: {name!r}")


def render_placeholder_slide(prs: Presentation, spec: dict):
    layout = _find_layout(prs, spec["layout"])
    slide = prs.slides.add_slide(layout)
    for idx_str, text in spec.get("fields", {}).items():
        idx = int(idx_str)
        ph = get_placeholder(slide, idx)
        if ph is not None and text:
            set_shape_text(ph, text)
    for idx_str, image_bytes in spec.get("images", {}).items():
        idx = int(idx_str)
        box = get_placeholder_box(layout, idx)
        ph = get_placeholder(slide, idx)
        if box and ph is not None:
            remove_shape_at(slide, list(slide.shapes).index(ph))
            insert_image_cover(slide, box, image_bytes)
    return slide


def _fill_table(shape, rows: list[list[str]]) -> None:
    table = shape.table
    n_rows = len(table.rows)
    n_cols = len(table.columns)
    for r in range(n_rows):
        for c in range(n_cols):
            cell = table.cell(r, c)
            if r < len(rows) and c < len(rows[r]):
                text = str(rows[r][c])
            else:
                text = ""
            set_shape_text_frame_only(cell.text_frame, text)


def set_shape_text_frame_only(tf, text: str) -> None:
    """Como set_shape_text pero operando directamente sobre un text_frame
    (para celdas de tabla, que no son 'shapes' con .text_frame homogéneo)."""
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    first_p = tf.paragraphs[0]
    for r in list(first_p.runs):
        r._r.getparent().remove(r._r)
    first_p.text = text


def render_duplicate_slide(prs: Presentation, spec: dict):
    """spec: {"example_index": int, "shape_text": {pos: text}, "images": {pos: bytes},
    "delete_positions": [pos, ...], "table": {"pos": int, "rows": [[..],..]}}"""
    slide = duplicate_slide(prs, spec["example_index"])
    # borra primero (de mayor a menor pos para no desordenar índices)
    for pos in sorted(spec.get("delete_positions", []), reverse=True):
        remove_shape_at(slide, pos)
    shapes = list(slide.shapes)
    for pos_str, text in spec.get("shape_text", {}).items():
        pos = int(pos_str)
        if 0 <= pos < len(shapes):
            set_shape_text(shapes[pos], text)
    for pos_str, image_bytes in spec.get("images", {}).items():
        pos = int(pos_str)
        if 0 <= pos < len(shapes):
            shp = shapes[pos]
            box = (shp.left, shp.top, shp.width, shp.height)
            remove_shape_at(slide, pos)
            insert_image_cover(slide, box, image_bytes)
    table_spec = spec.get("table")
    if table_spec:
        pos = int(table_spec["pos"])
        shapes = list(slide.shapes)
        if 0 <= pos < len(shapes) and shapes[pos].has_table:
            _fill_table(shapes[pos], table_spec["rows"])
    return slide


def build_deck(master_path: str, output_path: str, slide_specs: list[dict], example_master_path: Optional[str] = None):
    """example_master_path: si se pasa, se usa una copia SIN TOCAR de la plantilla
    (con sus 22 diapositivas de ejemplo intactas) como fuente para duplicar
    diapositivas de ejemplo, mientras que master_path es donde se construye el
    resultado final (que empieza vacío). Si no se pasa, se usa master_path para
    ambas cosas (duplicando antes de vaciar)."""
    prs = Presentation(master_path)

    if example_master_path:
        example_prs = Presentation(example_master_path)
    else:
        example_prs = prs

    rendered = []
    for spec in slide_specs:
        mode = spec["mode"]
        if mode == "placeholder":
            rendered.append(("placeholder", spec))
        elif mode == "duplicate":
            # duplicamos ya, desde la copia de ejemplo intacta, y guardamos el spec
            rendered.append(("duplicate", spec))
        else:
            raise ValueError(f"modo desconocido: {mode}")

    if example_master_path:
        # 1) construir usando prs "limpio" tras vaciar
        delete_all_slides(prs)
        for mode, spec in rendered:
            if mode == "placeholder":
                render_placeholder_slide(prs, spec)
            else:
                # duplicate needs slide from prs itself; copy example slide's XML
                # from example_prs into prs by duplicating within example_prs first,
                # then moving? Simpler: duplicate within prs is only valid if prs still
                # has the example slides. So when example_master_path is given, we
                # instead duplicate from example_prs into a temp, export shapes XML,
                # and graft into prs. To keep this simple & robust, require
                # example_master_path to be None for now (see build_deck_simple).
                raise NotImplementedError(
                    "usa build_deck sin example_master_path (más simple): "
                    "no borres las diapositivas antes de duplicar."
                )
    else:
        # enfoque simple y robusto: primero duplicamos TODO lo que haga falta
        # (mientras las 22 diapositivas de ejemplo siguen presentes al final),
        # y solo al terminar borramos las 22 originales.
        n_originals = len(prs.slides)
        for mode, spec in rendered:
            if mode == "placeholder":
                render_placeholder_slide(prs, spec)
            else:
                render_duplicate_slide(prs, spec)
        # borra las N diapositivas originales (las primeras n_originals)
        xml_slides = prs.slides._sldIdLst
        slide_ids = list(xml_slides)
        for sld in slide_ids[:n_originals]:
            rId = sld.get(qn("r:id"))
            prs.part.drop_rel(rId)
            xml_slides.remove(sld)

    prs.save(output_path)
    return output_path
