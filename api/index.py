"""
Entry point de Vercel (Python Serverless Function).

Vercel detecta cualquier módulo dentro de /api que exponga una app WSGI/ASGI
compatible. Aquí simplemente reexportamos la app Flask ya definida en
app.py (raíz del proyecto) para que Vercel la sirva.
"""
import sys
from pathlib import Path

# Permite importar app.py y engine/ desde la raíz del repo.
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import app  # noqa: E402

# Vercel's Python runtime looks for a WSGI callable named `app`.
