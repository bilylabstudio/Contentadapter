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
el hueco sin deformar).
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


def _prepare_cover_image(box, image_bytes: bytes) -> Optional[io.BytesIO]:
    """Prepara (recorta cover-fit) una imagen para llenar `box` (left, top,
    width, height en EMU) sin deformarla, SIN tocar la diapositiva. Devuelve
    None si la imagen no se puede decodificar (formato no soportado -- p.ej.
    WMF/EMF heredado de un .pptx antiguo -- bytes corruptos, dimensiones
    inválidas, etc.) en vez de propagar la excepción: así una única imagen
    problemática del documento de origen no puede tirar abajo la adaptación
    completa del archivo. El llamador debe comprobar el resultado antes de
    borrar la forma/placeholder original, para poder dejarlo intacto si la
    imagen no es utilizable (degradación grácil)."""
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
            # imagen más ancha de lo necesario -> recorta lados
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
            buf = _prepare_cover_image(box, image_bytes)
            if buf is not None:
                remove_shape_at(slide, list(slide.shapes).index(ph))
                slide.shapes.add_picture(buf, *box)
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
    for pos_str, image_bytes in spec.get("images", {}).items():
        pos = int(pos_str)
        if 0 <= pos < len(shapes):
            shp = shapes[pos]
            box = (shp.left, shp.top, shp.width, shp.height)
            buf = _prepare_cover_image(box, image_bytes)
            if buf is not None:
                remove_shape_at(slide, pos)
                slide.shapes.add_picture(buf, *box)
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


def make_qr_png(url: str) -> Optional[bytes]:
    """Genera un QR (PNG, fondo blanco, módulos oscuros) para `url`. Devuelve
    None si no hay URL o si la librería `segno` no está disponible / falla:
    el llamador degrada quitando el hueco de QR en vez de romper la
    presentación."""
    if not (url or "").strip():
        return None
    try:
        import segno
        qr = segno.make(url.strip(), error="m")
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
