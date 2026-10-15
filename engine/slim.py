"""
Aligera un .pptx pesado ANTES de abrirlo con python-pptx, procesando el zip
entrada a entrada para que el pico de memoria dependa de UNA sola imagen y no
del tamaño total del archivo.

Por qué: python-pptx carga en memoria todas las partes del paquete al abrirlo
(imágenes y vídeos incluidos). Un .pptx de 400 MB necesitaba más de 1 GB de RAM
solo para abrirse, en un servidor compartido con n8n y WhatsApp. La herramienta
solo usa el texto y unas pocas imágenes por diapositiva (que luego recorta y
reduce), así que no necesita la calidad original ni el vídeo/audio incrustado.

Reglas (solo sobre partes pesadas de ppt/media/ y ppt/embeddings/):
  - PNG/JPEG grandes  -> se reducen a max_px y se reescriben con el MISMO nombre
                         y formato (las relaciones del pptx no se tocan).
  - Otras imágenes muy grandes, vídeo, audio, objetos incrustados -> se vacían
    (parte con 0 bytes; python-pptx la abre igual y la tratamos como "sin imagen").
  - Todo lo demás (XML, layouts, rels, miniaturas pequeñas) se copia tal cual.
"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

from PIL import Image

BIG_BYTES = 400_000          # por debajo de esto una imagen se copia sin tocar
MAX_PX = 1600                # lado mayor tras reducir
JPEG_QUALITY = 82
KEEP_OTHER_IMAGE_BYTES = 2_000_000   # otras imágenes (gif/bmp/tiff/emf...) hasta 2 MB se copian
KEEP_EMBED_BYTES = 1_000_000          # objetos incrustados hasta 1 MB se copian

_SHRINKABLE = {".png", ".jpg", ".jpeg"}
_OTHER_IMAGES = {".gif", ".bmp", ".tif", ".tiff", ".webp", ".emf", ".wmf", ".svg"}

# Evita que una imagen gigante (descompresión) dispare la memoria: por encima de
# este número de píxeles se descarta en vez de decodificarse entera.
MAX_DECODE_PIXELS = 60_000_000


def _shrink(data: bytes, ext: str) -> bytes | None:
    """Reduce una imagen PNG/JPEG. Devuelve None si no se pudo (se vaciará)."""
    try:
        im = Image.open(io.BytesIO(data))
        w, h = im.size
        if w * h > MAX_DECODE_PIXELS:
            return None
        if ext in (".jpg", ".jpeg"):
            # draft() decodifica JPEG ya a escala reducida: mucha menos RAM
            try:
                im.draft("RGB", (MAX_PX, MAX_PX))
            except Exception:
                pass
        im.thumbnail((MAX_PX, MAX_PX))
        out = io.BytesIO()
        if ext == ".png":
            im.save(out, format="PNG", optimize=True)
        else:
            im.convert("RGB").save(out, format="JPEG", quality=JPEG_QUALITY)
        res = out.getvalue()
        return res if len(res) < len(data) else data
    except Exception:
        return None


def slim_pptx(src_path: str | Path, dst_path: str | Path) -> dict:
    """Escribe en dst_path una versión ligera de src_path. Devuelve estadísticas."""
    stats = {"entries": 0, "shrunk": 0, "emptied": 0, "bytes_in": 0, "bytes_out": 0}
    with zipfile.ZipFile(src_path) as zin, zipfile.ZipFile(
        dst_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6
    ) as zout:
        for info in zin.infolist():
            stats["entries"] += 1
            name = info.filename
            low = name.lower()
            ext = Path(low).suffix
            stats["bytes_in"] += info.file_size

            heavy_media = low.startswith("ppt/media/") and info.file_size > BIG_BYTES
            heavy_embed = low.startswith("ppt/embeddings/") and info.file_size > KEEP_EMBED_BYTES

            if not (heavy_media or heavy_embed):
                zout.writestr(info.filename, zin.read(info), compress_type=info.compress_type)
                continue

            if heavy_media and ext in _SHRINKABLE:
                new = _shrink(zin.read(info), ext)
                if new is not None:
                    zout.writestr(name, new)
                    stats["shrunk"] += 1
                    stats["bytes_out"] += len(new)
                    continue
            elif heavy_media and ext in _OTHER_IMAGES and info.file_size <= KEEP_OTHER_IMAGE_BYTES:
                zout.writestr(name, zin.read(info))
                stats["bytes_out"] += info.file_size
                continue

            # vídeo, audio, objetos incrustados o imagen imposible de reducir
            zout.writestr(name, b"")
            stats["emptied"] += 1
    return stats
