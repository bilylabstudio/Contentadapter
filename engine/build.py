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
real en su mismo bounding box. Por defecto se ajusta por 'cover' (se
recorta para llenar el hueco sin deformar), pero si la imagen de origen y
el hueco tienen proporciones muy distintas y ese recorte perdería más de un
30% de la imagen (p.ej. una captura de móvil muy vertical insertada en un
hueco apaisado), se usa 'contain' en su lugar: la imagen COMPLETA, sin
recortar, centrada sobre un fondo del color medio de su propio borde -- así
no se pierde texto, caras ni contenido importante cerca de los bordes de la
imagen original (bug detectado el 2026-11-18: capturas verticales
recortaban por arriba y por abajo, perdiendo cabeceras y caras).
"""
from __future__ import annotations

import copy
import io
import math
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


_EMU_PER_PT = 12700
_DEFAULT_FONT_PT = 18.0
_AVG_CHAR_WIDTH_FACTOR = 0.52  # ancho medio de carácter ~ 0.52 * tamaño de fuente (texto latino mixto)
_LINE_HEIGHT_FACTOR = 1.2      # alto de línea ~ 1.2 * tamaño de fuente (interlineado sencillo)
_MIN_FONT_SCALE = 0.5          # no reducir la fuente por debajo del 50% (deja de ser legible)
_MAX_LN_SPC_REDUCTION = 0.20   # tope habitual de PowerPoint para la reducción de interlineado


def _shape_text_box_pt(shape) -> Optional[tuple[float, float]]:
    """Ancho/alto en puntos del área de texto disponible dentro de la forma
    (dimensiones de la forma menos los márgenes internos del cuadro de
    texto, bIns/tIns/lIns/rIns -- usa los valores del XML si están
    explícitos, o los márgenes por defecto de PowerPoint en caso contrario)."""
    try:
        width_emu, height_emu = shape.width, shape.height
    except Exception:
        return None
    if not width_emu or not height_emu:
        return None
    bodyPr = shape.text_frame._txBody.find(qn("a:bodyPr"))

    def _inset(attr, default):
        if bodyPr is not None and bodyPr.get(attr) is not None:
            try:
                return int(bodyPr.get(attr))
            except ValueError:
                return default
        return default

    l_ins = _inset("lIns", 91440)
    r_ins = _inset("rIns", 91440)
    t_ins = _inset("tIns", 45720)
    b_ins = _inset("bIns", 45720)
    avail_w = (width_emu - l_ins - r_ins) / _EMU_PER_PT
    avail_h = (height_emu - t_ins - b_ins) / _EMU_PER_PT
    return max(avail_w, 1.0), max(avail_h, 1.0)

def _is_wide_char(ch: str) -> bool:
    """True para caracteres "anchos" (CJK, Hangul, kana, puntuación de ancho
    completo...) que ocupan aproximadamente un em completo en vez del ~0.5em
    tÍpico de una fuente latina. El contenido de IMSED es hoy en español,
    pero esto evita que el autoajuste calculado SUBESTIME el espacio
    necesario (y por tanto deje texto desbordado) si algún día se adapta
    contenido con texto en chino/japonés/coreano."""
    cp = ord(ch)
    return (
        0x1100 <= cp <= 0x11FF or  # Hangul Jamo
        0x3040 <= cp <= 0x30FF or  # Hiragana/Katakana
        0x3400 <= cp <= 0x4DBF or  # CJK Ext A
        0x4E00 <= cp <= 0x9FFF or  # CJK Unified
        0xAC00 <= cp <= 0xD7A3 or  # Hangul Syllables
        0xF900 <= cp <= 0xFAFF or  # CJK Compat
        0xFF00 <= cp <= 0xFFEF     # Full-width forms
    )


def _effective_char_width_factor(line: str) -> float:
    """Factor de ancho medio de carácter (como fracción del tamaño de
    fuente) para una línea concreta, ajustado al alza si una fracción
    relevante de sus caracteres son "anchos" (ver _is_wide_char)."""
    if not line:
        return _AVG_CHAR_WIDTH_FACTOR
    wide = sum(1 for ch in line if _is_wide_char(ch))
    frac_wide = wide / len(line)
    if frac_wide <= 0:
        return _AVG_CHAR_WIDTH_FACTOR
    # Un carácter ancho ocupa ~1.0em frente a ~0.52em de uno latino medio;
    # se interpola según la fracción de caracteres anchos de la línea.
    return _AVG_CHAR_WIDTH_FACTOR + frac_wide * (1.0 - _AVG_CHAR_WIDTH_FACTOR)


def _compute_autofit(shape, new_lines: list[str], rpr_templates: list) -> Optional[tuple[int, int]]:
    """Estima -- sin un motor de layout de texto real, que no está disponible
    aquí -- cuánto haría falta reducir la fuente (fontScale) y el
    interlineado (lnSpcReduction) para que el texto quepa en la caja, y
    devuelve ya los valores codificados como los espera <a:normAutofit>
    (milésimas de punto porcentual). Devuelve None si no hace falta reducir
    nada (el texto ya cabe a tamaño completo) o si no se puede estimar.

    Esto sustituye a dejar <a:normAutofit/> vacío confiando en que el propio
    visor (PowerPoint, Google Slides, LibreOffice...) recalcule el
    porcentaje de reducción en el momento de abrir el archivo -- ese cálculo
    "en vivo" no es fiable de forma consistente entre visores y es la causa
    más probable de las cajas de texto descuadradas reportadas por el
    cliente: con un valor explícito ya horneado en el archivo, todos los
    visores muestran lo mismo."""
    box = _shape_text_box_pt(shape)
    if box is None:
        return None
    avail_w, avail_h = box

    required_h = 0.0
    for i, line in enumerate(new_lines):
        t_idx = min(i, len(rpr_templates) - 1) if rpr_templates else -1
        rPr = rpr_templates[t_idx] if 0 <= t_idx < len(rpr_templates) else None
        font_pt = _DEFAULT_FONT_PT
        if rPr is not None and rPr.get("sz") is not None:
            try:
                font_pt = int(rPr.get("sz")) / 100.0
            except ValueError:
                pass
        char_w_pt = max(font_pt * _effective_char_width_factor(line), 1.0)
        line_h_pt = font_pt * _LINE_HEIGHT_FACTOR
        n_chars = max(len(line), 1)
        wrapped_lines = max(1, math.ceil((n_chars * char_w_pt) / avail_w))
        required_h += wrapped_lines * line_h_pt

    if required_h <= 0 or required_h <= avail_h:
        return None  # cabe a tamaño completo, no se toca nada

    scale = max(avail_h / required_h, _MIN_FONT_SCALE)
    if scale >= 0.999:
        return None

    font_scale_val = max(25000, min(100000, int(round(scale * 100)) * 1000))
    ln_spc_val = max(0, min(20000, int(round(min(_MAX_LN_SPC_REDUCTION, (1 - scale) * 0.6) * 100)) * 1000))
    return font_scale_val, ln_spc_val


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
    # las formas vecinas (título largo tapando el subtítulo, etc.).
    try:
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        # Lo anterior crea un <a:normAutofit/> vacío; python-pptx no calcula
        # ni escribe el % de reducción real, así que cada visor lo recalcula
        # "en vivo" a su manera (origen de cajas de texto descuadradas entre
        # PowerPoint/Google Slides/LibreOffice). Calculamos aquí una
        # estimación y la horneamos explícitamente en el XML para que el
        # resultado sea el mismo en cualquier visor.
        autofit_vals = _compute_autofit(shape, new_lines, rpr_templates)
        if autofit_vals is not None:
            bodyPr = txBody.find(qn("a:bodyPr"))
            if bodyPr is not None:
                normAutofit = bodyPr.find(qn("a:normAutofit"))
                if normAutofit is not None:
                    normAutofit.set("fontScale", str(autofit_vals[0]))
                    normAutofit.set("lnSpcReduction", str(autofit_vals[1]))
    except Exception:
        pass


def remove_shape_at(slide, pos: int) -> None:
    shapes = list(slide.shapes)
    if 0 <= pos < len(shapes):
        shp = shapes[pos]
        shp._element.getparent().remove(shp._element)

# Si el recorte 'cover' necesitaría descartar más de esta fracción de la
# imagen en el eje recortado, se usa 'contain' en su lugar (ver más abajo).
# Bug real detectado el 2026-11-18: con proporciones muy distintas (p.ej.
# una captura de móvil 9:16 en un hueco 4:3), el cover-fit anterior recortaba
# más de la mitad de la imagen, cortando cabeceras y caras.
_MAX_COVER_CROP_FRACTION = 0.30


def _edge_pad_color(im: "Image.Image", axis: str):
    """Color medio de una franja del borde de la imagen (superior/inferior si
    axis='h', izquierda/derecha si axis='v'), usado como fondo del margen al
    componer en modo 'contain'. Así el margen añadido se mezcla con la propia
    imagen en vez de dejar barras blancas que desentonen."""
    w, h = im.size
    try:
        if axis == "h":
            strip = im.crop((0, 0, w, max(1, h // 20)))
        else:
            strip = im.crop((0, 0, max(1, w // 20), h))
        return strip.resize((1, 1)).getpixel((0, 0))
    except Exception:
        return (255, 255, 255)


def _prepare_cover_image(box, image_bytes: bytes) -> Optional[io.BytesIO]:
    """Prepara una imagen para llenar `box` (left, top, width, height en EMU)
    sin deformarla, SIN tocar la diapositiva. Devuelve None si la imagen no
    se puede decodificar (formato no soportado -- p.ej. WMF/EMF heredado de
    un .pptx antiguo -- bytes corruptos, dimensiones inválidas, etc.) en vez
    de propagar la excepción: así una única imagen problemática del
    documento de origen no puede tirar abajo la adaptación completa del
    archivo. El llamador debe comprobar el resultado antes de borrar la
    forma/placeholder original, para poder dejarlo intacto si la imagen no
    es utilizable (degradación grácil).

    Por defecto usa 'cover' (recorta centrado para llenar el hueco). Si eso
    exigiría descartar más de `_MAX_COVER_CROP_FRACTION` de la imagen porque
    su proporción es muy distinta a la del hueco, usa 'contain' en su lugar:
    la imagen COMPLETA, sin recortar, centrada sobre un fondo del color medio
    de su propio borde -- para no perder contenido importante cerca de los
    bordes (texto, caras, UI...)."""
    left, top, width, height = box
    if not width or not height:
        return None
    try:
        im = Image.open(io.BytesIO(image_bytes))
        im = im.convert("RGB")
        src_w, src_h = im.size
        if src_w <= 0 or src_h <= 0:
            return None
        target_ratio = width / height
        src_ratio = src_w / src_h

        if src_ratio > target_ratio:
            # imagen más ancha de lo necesario -> recortaría los lados
            discarded = 1 - (max(1, int(src_h * target_ratio)) / src_w)
        else:
            # imagen más alta de lo necesario -> recortaría arriba/abajo
            discarded = 1 - (max(1, int(src_w / target_ratio)) / src_h)

        if discarded > _MAX_COVER_CROP_FRACTION:
            # 'contain': reescala la imagen completa para que quepa entera
            # dentro de la proporción del hueco, centrada sobre un fondo que
            # toma el color medio del borde de la propia imagen.
            work_w = 1600
            work_h = max(1, round(work_w * height / width))
            scale = min(work_w / src_w, work_h / src_h)
            new_w = max(1, round(src_w * scale))
            new_h = max(1, round(src_h * scale))
            resized = im.resize((new_w, new_h))
            pad_color = _edge_pad_color(im, "h" if src_ratio > target_ratio else "v")
            canvas = Image.new("RGB", (work_w, work_h), pad_color)
            canvas.paste(resized, ((work_w - new_w) // 2, (work_h - new_h) // 2))
            im = canvas
        else:
            if src_ratio > target_ratio:
                new_w = max(1, int(src_h * target_ratio))
                x0 = (src_w - new_w) // 2
                im = im.crop((x0, 0, x0 + new_w, src_h))
            else:
                new_h = max(1, int(src_w / target_ratio))
                y0 = (src_h - new_h) // 2
                im = im.crop((0, y0, src_w, y0 + new_h))

        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=88)
    except Exception:
        return None
    buf.seek(0)
    return buf


def insert_image_cover(slide, box, image_bytes: bytes) -> bool:
    """Inserta una imagen ocupando exactamente `box` (left, top, width, height
    en EMU), recortándola (cover-fit) para llenar el hueco sin deformarla.
    Devuelve True si se insertó, False si la imagen no se pudo decodificar
    (en cuyo caso no se ha añadido nada a la diapositiva)."""
    left, top, width, height = box
    buf = _prepare_cover_image(box, image_bytes)
    if buf is None:
        return False
    slide.shapes.add_picture(buf, left, top, width, height)
    return True


def _attach_hyperlink(pic, url: Optional[str]) -> None:
    """Si la imagen de origen (p.ej. la miniatura de un vídeo de
    TikTok/YouTube) tenía un hipervínculo asociado, lo añade también a la
    imagen insertada para que siga siendo "clicable" en la diapositiva
    adaptada. Antes de este fix la miniatura se copiaba como imagen plana y
    el enlace al vídeo se perdía por completo."""
    if not url:
        return
    try:
        pic.click_action.hyperlink.address = url
    except Exception:
        pass


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
    image_links = spec.get("image_links", {})
    for idx_str, image_bytes in spec.get("images", {}).items():
        idx = int(idx_str)
        box = get_placeholder_box(layout, idx)
        ph = get_placeholder(slide, idx)
        if box and ph is not None:
            buf = _prepare_cover_image(box, image_bytes)
            if buf is not None:
                remove_shape_at(slide, list(slide.shapes).index(ph))
                pic = slide.shapes.add_picture(buf, *box)
                _attach_hyperlink(pic, image_links.get(idx_str))
            # si la imagen no se pudo decodificar, se deja el placeholder de
            # texto/vacío original tal cual en vez de perder toda la diapositiva
    return slide

def _clear_cell_text(tc_el) -> None:
    """Vacía el texto de una celda de tabla clonada (de una fila/columna
    duplicada como plantilla), dejando su formato intacto (fuente, tamaño,
    alineación, bordes) pero sin repetir el texto de la fila/columna de
    origen."""
    txBody = tc_el.find(qn("a:txBody"))
    if txBody is None:
        return
    for p in txBody.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            t = r.find(qn("a:t"))
            if t is not None:
                t.text = ""


def _max_safe_table_rows(shape, sibling_shapes: list) -> int:
    """Calcula cuántas filas puede tener la tabla como máximo sin invadir la
    forma más próxima por debajo de ella en la diapositiva (p.ej. el pie de
    página con el número de diapositiva). Calcula el margen a partir de la
    geometría real de la diapositiva en vez de asumir un número de filas fijo,
    para que siga siendo correcto si la plantilla madre cambia."""
    table = shape.table
    current_rows = list(table.rows)
    if not current_rows:
        return 0
    avg_row_h = sum(r.height for r in current_rows) / len(current_rows)
    if not avg_row_h:
        return len(current_rows)
    table_bottom = shape.top + shape.height
    nearest_below = None
    for sib in sibling_shapes:
        if sib is shape or sib.top is None:
            continue
        if sib.top >= table_bottom and (nearest_below is None or sib.top < nearest_below):
            nearest_below = sib.top
    if nearest_below is None:
        # no hay ninguna forma debajo en esta diapositiva: usa el borde
        # inferior del slide (16:9, 7.5in) como límite conservador
        nearest_below = Emu(int(7.5 * 914400))
    headroom = nearest_below - table_bottom
    if headroom <= 0:
        return len(current_rows)
    extra_rows = int(headroom // avg_row_h)
    return len(current_rows) + max(0, extra_rows)


def _grow_table_rows(table, target_row_count: int) -> None:
    """Clona la última fila de la tabla (formato, bordes, alturas) hasta
    alcanzar target_row_count filas, vaciando el texto de las filas nuevas."""
    tbl = table._tbl
    trs = tbl.findall(qn("a:tr"))
    if not trs:
        return
    last_tr = trs[-1]
    while len(tbl.findall(qn("a:tr"))) < target_row_count:
        new_tr = copy.deepcopy(last_tr)
        for tc in new_tr.findall(qn("a:tc")):
            _clear_cell_text(tc)
        tbl.append(new_tr)
    table.notify_height_changed()


def _grow_table_cols(table, target_col_count: int) -> None:
    """Añade columnas a la tabla clonando la última columna (ancho y, en cada
    fila, formato de la última celda) y redistribuye el ancho total original
    a partes iguales entre todas las columnas, para que la tabla conserve su
    ancho total y no se salga de su hueco en la plantilla."""
    tbl = table._tbl
    tblGrid = tbl.find(qn("a:tblGrid"))
    if tblGrid is None:
        return
    grid_cols = tblGrid.findall(qn("a:gridCol"))
    n_cols = len(grid_cols)
    if n_cols == 0 or target_col_count <= n_cols:
        return
    total_width = sum(int(gc.get("w")) for gc in grid_cols)
    to_add = target_col_count - n_cols
    for _ in range(to_add):
        grid_cols = tblGrid.findall(qn("a:gridCol"))
        new_gc = copy.deepcopy(grid_cols[-1])
        tblGrid.append(new_gc)
        for tr in tbl.findall(qn("a:tr")):
            tcs = tr.findall(qn("a:tc"))
            if not tcs:
                continue
            new_tc = copy.deepcopy(tcs[-1])
            _clear_cell_text(new_tc)
            tr.append(new_tc)
    grid_cols = tblGrid.findall(qn("a:gridCol"))
    new_col_w = max(1, total_width // len(grid_cols))
    for gc in grid_cols:
        gc.set("w", str(new_col_w))
    table.notify_width_changed()


def _fill_table(shape, rows: list[list[str]], sibling_shapes: Optional[list] = None) -> None:
    """Rellena la tabla con `rows` (datos EXACTOS extraídos del documento de
    origen). Si esos datos no caben en la rejilla fija de la diapositiva de
    ejemplo (más filas y/o columnas de las que trae la plantilla), la hace
    crecer dinámicamente en vez de truncar en silencio -- el crecimiento de
    filas se limita al margen de seguridad real antes de la siguiente forma
    de la diapositiva (p.ej. el pie de página), para no pisar ningún
    elemento; si aun así hubiera más filas de las que caben con margen, se
    usa ese máximo seguro como último recurso (en vez de solapar formas)."""
    table = shape.table
    n_rows = len(table.rows)
    n_cols = len(table.columns)

    needed_rows = len(rows)
    needed_cols = max((len(r) for r in rows), default=0)

    if needed_rows > n_rows and sibling_shapes is not None:
        safe_max_rows = _max_safe_table_rows(shape, sibling_shapes)
        target_rows = min(needed_rows, safe_max_rows)
        if target_rows > n_rows:
            _grow_table_rows(table, target_rows)
            n_rows = len(table.rows)

    if needed_cols > n_cols:
        _grow_table_cols(table, needed_cols)
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
    image_links = spec.get("image_links", {})
    for pos_str, image_bytes in spec.get("images", {}).items():
        pos = int(pos_str)
        if 0 <= pos < len(shapes):
            shp = shapes[pos]
            box = (shp.left, shp.top, shp.width, shp.height)
            buf = _prepare_cover_image(box, image_bytes)
            if buf is not None:
                remove_shape_at(slide, pos)
                pic = slide.shapes.add_picture(buf, *box)
                _attach_hyperlink(pic, image_links.get(pos_str))
            # si la imagen no se pudo decodificar, se deja la forma original
            # (p.ej. el placeholder de imagen de ejemplo) tal cual, en vez de
            # perder toda la diapositiva
    table_spec = spec.get("table")
    if table_spec:
        pos = int(table_spec["pos"])
        shapes = list(slide.shapes)
        if 0 <= pos < len(shapes) and shapes[pos].has_table:
            _fill_table(shapes[pos], table_spec["rows"], sibling_shapes=shapes)
    return slide

# ---------------------------------------------------------------------------
# Diapositiva de cierre con QR de valoración por máster
# ---------------------------------------------------------------------------

CLOSING_EXAMPLE_INDEX = 21  # diapositiva "Muchas gracias" + tarjeta con hueco de QR
CLOSING_LAYOUT = "CIERRE"
MAX_VISIBLE_LINK_CHARS = 28  # ancho útil del texto del enlace en la tarjeta de valoración


def make_qr_png(url: str) -> Optional[bytes]:
    """Genera un QR (PNG, fondo blanco, módulos oscuros) para `url`. Devuelve
    None si no hay URL o si la librería `segno` no está disponible / falla:
    el llamador degrada quitando el hueco de QR en vez de romper la
    presentación."""
    if not (url or "").strip():
        return None
    try:
        import segno
        u = url.strip()
        # URLs largas (Google Forms): corrección "L" -> menos módulos, más grandes
        # y fáciles de escanear en el hueco de 1,45 in de la tarjeta
        qr = segno.make(u, error="l" if len(u) > 60 else "m")
        buf = io.BytesIO()
        qr.save(buf, kind="png", scale=12, border=1, dark="#1c2f57", light="#ffffff")
        return buf.getvalue()
    except Exception:
        return None


def _find_qr_shapes(slide):
    """Devuelve (cuadrado_hueco, etiqueta_QR) de la diapositiva de cierre. Se
    localizan por contenido/geometría (etiqueta con texto 'QR' y el cuadrado
    que la contiene), con las posiciones conocidas de la plantilla como
    respaldo."""
    shapes = list(slide.shapes)
    label = None
    for shp in shapes:
        if shp.has_text_frame and shp.text_frame.text.strip().upper() == "QR":
            label = shp
            break
    square = None
    if label is not None:
        for shp in shapes:
            if shp is label or not shp.width or not shp.height:
                continue
            near_square = abs(shp.width - shp.height) <= 0.08 * max(shp.width, shp.height)
            contains = (shp.left <= label.left and shp.top <= label.top
                        and shp.left + shp.width >= label.left + label.width
                        and shp.top + shp.height >= label.top + label.height)
            if near_square and contains:
                square = shp
                break
    if square is None and len(shapes) > 4:
        square = shapes[3]
    if label is None and len(shapes) > 4 and shapes[4].has_text_frame:
        label = shapes[4]
    return square, label


def _display_url(url: str) -> str:
    u = (url or "").strip()
    for prefix in ("https://", "http://"):
        if u.lower().startswith(prefix):
            u = u[len(prefix):]
    if u.lower().startswith("www."):
        u = u[4:]
    return u.rstrip("/")


def _sync_visible_link(slide, url: str) -> None:
    """El texto visible 'imsed.com/valora' de la tarjeta debe coincidir con el
    destino real del QR: si el máster tiene su propio formulario, se muestra
    esa dirección (sin esquema)."""
    shown = _display_url(url)
    if not shown:
        return
    if len(shown) > MAX_VISIBLE_LINK_CHARS:
        # enlaces largos (p.ej. Google Forms) no caben en la tarjeta: se
        # sustituyen por un texto corto; el destino real va en el QR.
        shown = "Escanea el código QR"
    for shp in slide.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip().lower() == "imsed.com/valora":
            if shown.lower() != "imsed.com/valora":
                set_shape_text(shp, shown)
            return


def place_qr_on_closing(slide, qr_png: Optional[bytes], url: Optional[str] = None) -> bool:
    """Sustituye el hueco 'QR' de la diapositiva de cierre por la imagen del
    QR (mismo tamaño y posición). Si no hay QR utilizable, se QUITA el hueco
    y su etiqueta para no dejar un cuadro vacío con la palabra 'QR'. Devuelve
    True si se insertó un QR."""
    if url and qr_png:
        _sync_visible_link(slide, url)
    square, label = _find_qr_shapes(slide)
    if square is None:
        return False
    box = (square.left, square.top, square.width, square.height)
    ok = False
    if qr_png:
        try:
            Image.open(io.BytesIO(qr_png)).verify()
            ok = True
        except Exception:
            ok = False
    for shp in (square, label):
        if shp is not None:
            shp._element.getparent().remove(shp._element)
    if ok:
        slide.shapes.add_picture(io.BytesIO(qr_png), *box)
    return ok


def normalize_closing_spec(slide_specs: list[dict]) -> list[dict]:
    """Garantiza que la presentación termina con UNA sola diapositiva de
    cierre (con tarjeta de valoración + hueco de QR): quita cualquier cierre
    que haya propuesto el mapeo (conservando su texto de título, si lo
    había) y añade uno al final."""
    title = None
    kept = []
    for spec in slide_specs:
        is_closing = (
            (spec.get("mode") == "duplicate" and spec.get("example_index") == CLOSING_EXAMPLE_INDEX)
            or (spec.get("mode") == "placeholder" and spec.get("layout") == CLOSING_LAYOUT)
        )
        if is_closing:
            if title is None:
                if spec.get("mode") == "duplicate":
                    title = (spec.get("shape_text") or {}).get("0")
                else:
                    title = (spec.get("fields") or {}).get("0")
            continue
        kept.append(spec)
    closing = {
        "mode": "duplicate",
        "example_index": CLOSING_EXAMPLE_INDEX,
        "shape_text": {"0": title} if title else {},
        "images": {},
        "delete_positions": [],
        "_closing": True,
    }
    kept.append(closing)
    return kept

def build_deck(master_path: str, output_path: str, slide_specs: list[dict], example_master_path: Optional[str] = None,
               qr_url: Optional[str] = None):
    """example_master_path: si se pasa, se usa una copia SIN TOCAR de la plantilla
    (con sus 22 diapositivas de ejemplo intactas) como fuente para duplicar
    diapositivas de ejemplo, mientras que master_path es donde se construye el
    resultado final (que empieza vacío). Si no se pasa, se usa master_path para
    ambas cosas (duplicando antes de vaciar).

    qr_url: si se pasa, la presentación termina SIEMPRE con una única
    diapositiva de cierre cuyo hueco de QR contiene el QR de esa URL
    (formulario de valoración del máster). Si no se pasa, no se toca el
    cierre."""
    prs = Presentation(master_path)
    if qr_url is not None:
        slide_specs = normalize_closing_spec(slide_specs)
        qr_png = make_qr_png(qr_url)

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
                new_slide = render_duplicate_slide(prs, spec)
                if spec.get("_closing"):
                    place_qr_on_closing(new_slide, qr_png, qr_url)
        # borra las N diapositivas originales (las primeras n_originals)
        xml_slides = prs.slides._sldIdLst
        slide_ids = list(xml_slides)
        for sld in slide_ids[:n_originals]:
            rId = sld.get(qn("r:id"))
            prs.part.drop_rel(rId)
            xml_slides.remove(sld)

    prs.save(output_path)
    return output_path
