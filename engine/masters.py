"""
Registro de másters IMSED -> URL del formulario de valoración (Google Forms).

La URL de cada máster se convierte en un QR que se inserta en la última
diapositiva (cierre) de la presentación adaptada. El registro vive en
`masters.json` (raíz del repo): añadir un máster nuevo es añadir una entrada,
sin tocar código. Un máster sin `forms_url` usa la URL del máster por
defecto (`default`), para que la diapositiva de cierre siempre lleve un QR
válido.

`resolve_master()` acepta tanto el `id` como el nombre del máster, ignorando
mayúsculas, tildes y signos; así el flujo de Drive puede pasar directamente el
nombre de la subcarpeta como valor del máster.
"""
from __future__ import annotations

import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Optional

MASTERS_PATH = Path(os.environ.get("MASTERS_JSON", Path(__file__).resolve().parent.parent / "masters.json"))


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def _load() -> dict:
    with open(MASTERS_PATH, encoding="utf-8") as f:
        data = json.load(f)
    masters = data.get("masters") or []
    if not masters:
        raise ValueError("masters.json no contiene ningún máster")
    default_id = data.get("default") or masters[0]["id"]
    return {"default": default_id, "masters": masters}


def list_masters() -> list[dict]:
    data = _load()
    return [
        {"id": m["id"], "name": m["name"], "has_forms_url": bool((m.get("forms_url") or "").strip())}
        for m in data["masters"]
    ]


def masters_without_url() -> list[str]:
    return [m["id"] for m in list_masters() if not m["has_forms_url"]]


def default_master() -> dict:
    data = _load()
    return next(m for m in data["masters"] if m["id"] == data["default"])


def resolve_master(value: Optional[str]) -> Optional[dict]:
    """Devuelve el máster correspondiente a `value` (id o nombre, tolerante a
    tildes/mayúsculas/prefijos como 'Máster en'), el máster por defecto si
    `value` está vacío, o None si no coincide con ninguno."""
    data = _load()
    if not (value or "").strip():
        return default_master()
    target = _norm(value)
    for m in data["masters"]:
        if _norm(m["id"]) == target or _norm(m["name"]) == target:
            return m
    # coincidencia parcial (p.ej. nombre de carpeta "Martech" o "Master en Martech - ...")
    candidates = [
        m for m in data["masters"]
        if m["id"] != data["default"] and (
            _norm(m["id"]) in target
            or target in _norm(m["name"])
            or _norm(m["name"]) in target
        )
    ]
    return candidates[0] if len(candidates) == 1 else None


def forms_url_for(master: dict) -> str:
    """URL del formulario del máster; si no la tiene, la del máster por defecto."""
    url = (master.get("forms_url") or "").strip()
    return url or (default_master().get("forms_url") or "").strip()
