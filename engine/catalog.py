"""
Catálogo de plantillas de diapositiva disponibles en la plantilla madre IMSED
(assets/master_template.pptx). Cada entrada describe UNA diapositiva de
ejemplo concreta de esa plantilla (identificada por su índice original,
`example_index`) que el motor de renderizado duplicará y rellenará.

Esto es lo que se le pasa al LLM (mapper.py) como "menú" de diseños entre
los que elegir para cada diapositiva del PPT/PDF antiguo. Los índices de
posición (`pos`) son estables: corresponden al orden de `slide.shapes` de
esa diapositiva de ejemplo dentro de master_template.pptx, y son los mismos
que usa `engine/build.py` para `shape_text` / `images` / `delete_positions`.

Si en el futuro cambia la plantilla madre, este catálogo hay que
regenerarlo/ajustarlo a mano (ver scripts/inspect_master.py).
"""

TEMPLATES = [
    {
        "id": "portada",
        "example_index": 1,
        "layout": "PORTADA",
        "when": "Primera diapositiva de la presentación: título de la clase/sesión, máster y docente.",
        "editable": {
            "0": "kicker -- nombre del máster (mayúsculas, corto)",
            "1": "título -- título de la clase o sesión (máx. 2 líneas)",
            "2": "info -- 2-3 líneas cortas: módulo/sesión, docente, fecha (usa \\n entre líneas)",
        },
    },
    {
        "id": "seccion",
        "example_index": 5,
        "layout": "SECCION",
        "when": "Diapositiva divisoria que abre un bloque/tema nuevo dentro de la presentación.",
        "editable": {
            "0": "número de bloque, formato '/01', '/02'... (o '/?' si no aplica)",
            "1": "título de la sección (corto, 1-2 líneas)",
            "2": "subtítulo -- qué se va a ver en este bloque (1 línea)",
        },
    },
    {
        "id": "contenido_bullets",
        "example_index": 4,
        "layout": "CONTENIDO",
        "when": "Diapositiva de texto con título + subtítulo + lista de puntos (agenda, objetivos, bibliografía, cualquier lista con viñetas). Usa \\n para separar cada punto en el campo 3.",
        "editable": {
            "0": "kicker corto (contexto: módulo/sesión o similar)",
            "1": "título de la diapositiva",
            "2": "subtítulo o frase guía (1 línea)",
            "3": "cuerpo -- lista de puntos separados por \\n (3 a 6 puntos, una idea por línea)",
        },
    },
    {
        "id": "contenido_simple",
        "example_index": 6,
        "layout": "CONTENIDO",
        "when": "Diapositiva de texto con título + subtítulo + UN solo bloque de desarrollo (no es una lista, es una idea desarrollada en un párrafo corto o 2-3 frases).",
        "editable": {
            "0": "kicker corto",
            "1": "título de la diapositiva",
            "2": "subtítulo o pregunta guía",
            "3": "desarrollo -- una idea, en frases cortas (máx. 5 líneas)",
        },
    },
    {
        "id": "dato_destacado",
        "example_index": 12,
        "layout": "DATO DESTACADO",
        "when": "Diapositiva de UN dato/estadística grande y llamativa (una cifra, un porcentaje) con su explicación.",
        "editable": {
            "0": "kicker, ej. 'EL DATO DE LA SESIÓN' o similar",
            "1": "el dato en grande, ej. '73 %' o '2,4x' (muy corto)",
            "2": "explicación -- qué significa, fuente y periodo (2-3 líneas)",
        },
    },
    {
        "id": "texto_imagen",
        "example_index": 7,
        "layout": "TEXTO + IMAGEN",
        "when": "Texto a la izquierda + una imagen real a la derecha que aporta información (foto, captura, gráfico).",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "3": "texto -- 2-4 frases cortas o puntos separados por \\n",
        },
        "image_slot": "6",
        "delete_if_no_image": ["4", "5"],
    },
    {
        "id": "imagen_completa",
        "example_index": 17,
        "layout": "IMAGEN COMPLETA",
        "when": "Una imagen a sangre completa (ocupa toda la diapositiva), con un pie de foto pequeño. Úsalo para fotos/collages/capturas que no tienen texto que explicar, solo ilustran.",
        "editable": {
            "2": "pie de imagen -- fuente y año, o descripción muy corta",
        },
        "image_slot": "3",
        "delete_if_no_image": ["0", "1"],
    },
    {
        "id": "tarjetas_3",
        "example_index": 8,
        "layout": "TARJETAS",
        "when": "Tres ideas/bloques del mismo peso, en paralelo (columnas). Usa esto para listas de 3 conceptos, 3 factores, 3 tipos, etc.",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "5": "título tarjeta 1 (corto)",
            "6": "cuerpo tarjeta 1 (2-3 líneas)",
            "9": "título tarjeta 2 (corto)",
            "10": "cuerpo tarjeta 2 (2-3 líneas)",
            "13": "título tarjeta 3 (corto)",
            "14": "cuerpo tarjeta 3 (2-3 líneas)",
        },
        "groups": [["3", "4", "5", "6"], ["7", "8", "9", "10"], ["11", "12", "13", "14"]],
        "min_items": 2,
        "max_items": 3,
    },
    {
        "id": "tarjetas_4",
        "example_index": 9,
        "layout": "TARJETAS",
        "when": "Cuatro ideas/bloques cortos del mismo peso (etiqueta + frase). Úsalo para 4 elementos, no 3.",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "4": "etiqueta tarjeta 1 (1-2 palabras)",
            "5": "texto tarjeta 1 (1 línea)",
            "7": "etiqueta tarjeta 2",
            "8": "texto tarjeta 2",
            "10": "etiqueta tarjeta 3",
            "11": "texto tarjeta 3",
            "13": "etiqueta tarjeta 4",
            "14": "texto tarjeta 4",
        },
        "groups": [["3", "4", "5"], ["6", "7", "8"], ["9", "10", "11"], ["12", "13", "14"]],
        "min_items": 4,
        "max_items": 4,
    },
    {
        "id": "tarjetas_terminos",
        "example_index": 14,
        "layout": "TARJETAS",
        "when": "Glosario / lista de términos clave, cada uno con una definición corta de una línea. Hasta 6 términos.",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "3": "término 1", "4": "definición 1 (1 línea)",
            "6": "término 2", "7": "definición 2 (1 línea)",
            "9": "término 3", "10": "definición 3 (1 línea)",
            "12": "término 4", "13": "definición 4 (1 línea)",
            "15": "término 5", "16": "definición 5 (1 línea)",
            "18": "término 6", "19": "definición 6 (1 línea)",
        },
        "groups": [["3", "4", "5"], ["6", "7", "8"], ["9", "10", "11"], ["12", "13", "14"], ["15", "16", "17"], ["18", "19", "20"]],
        "min_items": 2,
        "max_items": 6,
    },
    {
        "id": "tarjetas_takeaways",
        "example_index": 18,
        "layout": "TARJETAS",
        "when": "Cierre de bloque tipo 'tres ideas para llevarse' -- resumen de 3 conclusiones o aprendizajes clave, numeradas, sin caja de fondo.",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "4": "idea 1 (frase completa)", "5": "detalle idea 1 (1 línea, opcional)",
            "7": "idea 2 (frase completa)", "8": "detalle idea 2 (1 línea, opcional)",
            "10": "idea 3 (frase completa)", "11": "detalle idea 3 (1 línea, opcional)",
        },
        "groups": [["3", "4", "5"], ["6", "7", "8"], ["9", "10", "11"]],
        "min_items": 2,
        "max_items": 3,
    },
    {
        "id": "pasos",
        "example_index": 10,
        "layout": "PASOS",
        "when": "Secuencia, proceso o cronología de pasos ordenados (2 a 4 pasos).",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "6": "título paso 1", "7": "descripción paso 1 (1 línea)",
            "10": "título paso 2", "11": "descripción paso 2 (1 línea)",
            "14": "título paso 3", "15": "descripción paso 3 (1 línea)",
            "18": "título paso 4", "19": "descripción paso 4 (1 línea)",
        },
        "groups": [["4", "5", "6", "7"], ["8", "9", "10", "11"], ["12", "13", "14", "15"], ["16", "17", "18", "19"]],
        "min_items": 2,
        "max_items": 4,
    },
    {
        "id": "dos_columnas",
        "example_index": 11,
        "layout": "DOS COLUMNAS",
        "when": "Comparativa de dos alternativas, antes/después, pros/contras, opción A vs opción B.",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "4": "etiqueta columna izquierda (ej. 'ANTES', 'OPCIÓN A')",
            "5": "puntos columna izquierda, separados por \\n (2-4 puntos)",
            "7": "etiqueta columna derecha (ej. 'DESPUÉS', 'OPCIÓN B')",
            "8": "puntos columna derecha, separados por \\n (2-4 puntos)",
        },
    },
    {
        "id": "tabla",
        "example_index": 13,
        "layout": "TABLA",
        "when": "Datos tabulares reales (filas x columnas) que aparecían como tabla en el original. Máximo 5 filas de datos (más una fila de cabecera).",
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
        },
        "table_slot": "3",
    },
    {
        "id": "ejercicio",
        "example_index": 16,
        "layout": "EJERCICIO",
        "when": "Caso práctico / actividad / ejercicio para los alumnos, con encargo y modo de trabajo.",
        "editable": {
            "0": "kicker corto",
            "1": "título del ejercicio",
            "2": "modalidad, ej. 'Trabajo en grupo', 'Individual'",
            "3": "EL ENCARGO\\n<qué hay que hacer>\\n \\nCÓMO SE TRABAJA\\n<cómo>\\n \\nQUÉ SE ENTREGA\\n<entregable>",
            "5": "duración, ej. '30 min' (opcional, deja el original si no se sabe)",
        },
    },
    {
        "id": "cierre",
        "example_index": 21,
        "layout": "CIERRE",
        "when": "Última diapositiva de agradecimiento/cierre.",
        "editable": {
            "0": "texto de cierre, normalmente 'Muchas gracias'",
        },
    },
    {
        "id": "imagen_texto",
        "example_index": None,
        "layout": "IMAGEN + TEXTO",
        "when": "Como texto_imagen pero con la imagen a la IZQUIERDA y el texto a la derecha.",
        "placeholder_layout": True,
        "editable": {
            "0": "kicker corto",
            "1": "título",
            "2": "subtítulo (1 línea)",
            "4": "texto -- 2-4 frases cortas o puntos separados por \\n",
        },
        "image_field": "3",
    },
]

BY_ID = {t["id"]: t for t in TEMPLATES}


def catalog_prompt_text() -> str:
    """Serializa el catálogo a texto compacto para el prompt del LLM."""
    lines = []
    for t in TEMPLATES:
        lines.append(f"### {t['id']}  (layout: {t['layout']})")
        lines.append(f"Cuándo usarla: {t['when']}")
        if t.get("groups"):
            lines.append(f"Es una plantilla de tarjetas/bloques repetidos: mínimo {t['min_items']}, máximo {t['max_items']} unidades. Si tienes menos, omite las unidades sobrantes (vía 'omit_groups').")
        if t.get("image_slot") or t.get("image_field"):
            lines.append("Esta plantilla admite UNA imagen real (campo 'image').")
        if t.get("table_slot"):
            lines.append("Esta plantilla tiene una TABLA real: rellena 'table_rows' (lista de filas, cada fila lista de celdas de texto; primera fila = cabecera; máx. 6 filas incl. cabecera).")
        lines.append("Campos editables (posición -> qué va ahí):")
        for pos, desc in t["editable"].items():
            lines.append(f"  {pos}: {desc}")
        lines.append("")
    return "\n".join(lines)
