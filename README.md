# IMSED · Adaptador de presentaciones

Sube una o varias presentaciones antiguas (`.pptx` o `.pdf`) y las devuelve
adaptadas al diseño de la plantilla madre de IMSED, conservando el
contenido pero reconstruyéndolo con las diapositivas de ejemplo,
tipografías y colores de la plantilla nueva.

Stateless: no hay base de datos ni almacenamiento persistente. Cada request
procesa sus archivos en memoria/temporal y los descarta al responder. La API
key de Gemini la manda el propio navegador en cada llamada (no se guarda en
el servidor).

## Cómo funciona

1. `engine/extract.py` -- saca el texto y las imágenes reales de cada
   diapositiva/página del documento antiguo (pptx con python-pptx, pdf con
   pypdf).
2. `engine/mapper.py` -- por cada diapositiva de origen, llama a Gemini con
   el catálogo de plantillas disponibles (`engine/catalog.py`) y le pide que
   elija la que mejor encaja y redacte el texto para cada hueco.
3. `engine/build.py` -- duplica la diapositiva de ejemplo elegida dentro de
   la plantilla madre y sustituye su texto/imágenes, preservando el diseño
   (colores, viñetas, tarjetas, etc.) tal cual está en la plantilla.

## Desplegar en Vercel

Este repo está preparado como función serverless de Python para Vercel
(mismo patrón que CV Builder IMSED: repo en GitHub + proyecto Vercel
conectado a ese repo, deploy automático en cada push a `main`).

- `api/index.py` -- reexporta la app Flask de `app.py` como función WSGI.
- `vercel.json` -- enruta todo el tráfico a esa función y le da hasta 300s
  y 1024MB de memoria (los batches con Gemini + renderizado de varias
  diapositivas pueden tardar).
- No hace falta ninguna base de datos ni volumen persistente, ni variables
  de entorno obligatorias -- la API key de Gemini la introduce cada usuario
  en el navegador.

Pasos: crear un proyecto en Vercel apuntando a este repo (framework preset
"Other"), deploy. No hay build step ni variables que configurar.

**Aviso importante:** las Serverless Functions de Vercel limitan el tamaño
del body de la request (por defecto 4.5MB en el plan Hobby). Los límites de
la app (25 archivos, 60MB por archivo) están pensados para un servidor
tradicional (EasyPanel/Docker); en Vercel, lotes o archivos grandes pueden
devolver 413 antes de llegar a Flask. Si IMSED va a subir presentaciones
grandes habitualmente, conviene revisar el límite de payload del plan de
Vercel o mantener también la opción de EasyPanel como fallback.

## La plantilla madre en este repo (Vercel)

Este repo NO incluye el binario `master_template.pptx` (Vercel no tiene
disco persistente para bundlearlo cómodamente). En su lugar, `app.py`
lo descarga la primera vez desde la URL pública de exportación de Google
Slides del documento original de IMSED y lo cachea en `/tmp` para las
siguientes invocaciones del mismo contenedor. Si IMSED cambia la plantilla
madre, hay que actualizar `MASTER_TEMPLATE_URL` en `app.py` (o volver a
compartir el nuevo Google Slides como "cualquiera con el enlace, editor")
Y actualizar `engine/catalog.py` (los índices de diapositiva y posiciones
de forma son específicos de la plantilla actual).

## Límites conocidos (v1)

- Pensado para el estilo de máster/sesión de IMSED (14 tipos de diapositiva
  de la plantilla actual). Contenido muy atípico (SmartArt complejo,
  infografías) se simplifica a texto o a una imagen a sangre completa.
- Tablas: se recrean como tabla real solo hasta 5 filas + cabecera (el
  límite que marca la propia plantilla).
- El PDF de origen debe tener texto seleccionable (no imágenes escaneadas
  sin OCR); si una página es 100% imagen sin texto, se traduce a una
  diapositiva de "imagen completa" sin texto adicional.
- Máximo 25 archivos y 60MB por archivo en un mismo lote.
