# Referencia de API — engine/mb_engine.js y módulos relacionados

Esta es la firma completa de cada función exportada. Úsala para no tener
que adivinar parámetros ni releer el código fuente cada vez.

## mb_engine.js

### deriveGrid(opts) → grid

Calcula el grid completo para una slide de contenido. Llamar SIEMPRE
primero, antes de cualquier otro helper.

```js
const grid = E.deriveGrid({
  bodySize: 11,
  leadingRatio: 1.4,
  marginL: 0.40,
  marginR: 0.40,
  columns: "asym32",
  footerH: 0.28,
  hasNotes: false,
  hasHeader: false,
});
```

Devuelve un objeto con slideW, slideH, margins, module, gutter, workW,
cols (L/R con x y w), y zones (contentTop, contentBottom, footerY, etc).

Valores válidos de columns: "sym2" (dos columnas iguales), "asym32"
(60/40), "single" (una sola columna a todo el ancho).

### deriveHeaderBlock(grid, titleText, opts) → head

titleText es obligatorio — se usa para estimar cuántas líneas ocupará
el título real y ajustar tamaño de fuente, evitando desborde sobre el
subtítulo. opts acepta titleSize (default 22) y gapAfterSubtitle
(default 0.9). Devuelve ruleY, ruleH, titleY, titleH, subY, subH,
blocksY, titleSize, titleLines. blocksY es el Y donde debe empezar el
contenido real.

### addHeader(slide, pres, grid, head, opts)

Dibuja línea dorada + título + subtítulo + logo opcional. opts acepta
title (debe coincidir con el pasado a deriveHeaderBlock), subtitle (o
null), logoPath opcional, logoH opcional (default 0.5).

### addFooter(slide, pres, grid, opts)

opts acepta left, center, right — los tres segmentos de texto del
footer navy.

### addNotes(slide, grid, notesArray)

Solo dibuja si grid.zones.hasNotes es true. notesArray es un array de
strings, cada uno una línea de nota.

### tblCell(text, opts) → objeto celda para slide.addTable()

opts acepta fs (fontSize), bold, ul (underline), color, align, fill.

### inferProjectName(title) → string

Infiere el nombre corto de proyecto para el footer a partir de un
título de slide. Heurística simple — verificar visualmente el
resultado con títulos atípicos.

### addCard(slide, pres, card)

Tarjeta de color con título + cuerpo, con protección de desborde en
ambos. card acepta x, y, w, h, color, title, body, titleSize (default
11), bodySize (default 7.5), module (opcional, pasar grid.module).

### layoutFlowchart(steps, availableH, boxW, opts) → layout

Calcula las alturas de N cajas de flowchart verticales para que el
total nunca exceda availableH. steps es un array de {label, body}.
opts acepta labelSize, bodySize, arrowH (todos opcionales).

Devuelve steps (con h calculado por cada uno), labelSize, bodySize,
arrowH finales, y overflow/overflowAmount/totalNeeded/availableH.

Si overflow es true, NO dibujar todas las cajas con el h calculado de
todos modos — eso produciría cajas que rebasan contentBottom. Dibujar
solo los pasos que caben (acumulando cursorY y comparando contra
grid.zones.contentBottom antes de cada uno) y avisar al usuario. Ver
el patrón completo en tests/stress_test_full.js, caso
C3_flowchart_9_pasos_overflow_esperado.

### addFlowStep(slide, pres, opts)

Dibuja UNA caja de flowchart ya dimensionada (no calcula tamaños — eso
lo hace layoutFlowchart). Llamar una vez por cada paso, acumulando
manualmente la posición Y. opts acepta x, y, w, h, label, body, isEdge
(true = caja navy para primer/último paso), labelSize, bodySize,
drawArrowBelow, arrowH.

## mb_cover.js

### buildCover(slide, pres, opts)

Template FIJO de portada CV. No deriva del grid de contenido — es
identidad de marca, siempre la misma estructura.

opts acepta company (supratítulo dorado), title, subtitle, client
(aparece en header "PREPARADO PARA X"), month, city, country,
heroImagePath (null para usar placeholder), logoCvPath.

## mb_text_fit.js

### fitTextLines(text, widthIn, fontSize, maxLines) → {fontSize, lines}

Función de bajo nivel — normalmente no se llama directamente, la usan
deriveHeaderBlock y addCard internamente. Solo se necesita si se está
construyendo un layout nuevo que requiere la misma protección.

CHAR_WIDTH_RATIO de 0.58 está calibrado para texto BOLD en
Arial/Cambria en español/inglés. El reporte de lines es honesto: si
tras reducir al piso (mínimo entre 6 y el fontSize solicitado) el
texto sigue ocupando más de maxLines, lo dice — no lo trunca
artificialmente. El caller debe manejar ese caso.

## mb_images.js

### fetchHeroImage(query, outDir, opts) → Promise con path, photographer, photographerUrl, sourceId

Busca en Pexels, descarga la mejor opción, recorta al aspect ratio del
hero de portada (10:2.95 por default), y devuelve la ruta local lista
para usar en buildCover con heroImagePath igual al path devuelto.

Query en inglés suele dar mejores resultados que en español. opts
acepta filenameHint, aspectW, aspectH (todos opcionales).

Requiere config/api_keys.json con una key de Pexels válida. Lanza
error si no existe o si la búsqueda no devuelve resultados.

## mb_validator.js

CLI: node engine/mb_validator.js archivo.pptx — correr después de
generar cualquier archivo, antes de entregarlo. Exit code 0 si pasa, 1
si encuentra problemas.

Programático: validatePptx(pptxPath, boundsPerSlide) devuelve un array
de issues. Longitud cero significa que pasó. Detecta FOOTER_OVERLAP,
NOTES_ZONE_OVERLAP, OUT_OF_BOUNDS, TEXT_OVERLAP (excluyendo el patrón
intencional de texto-sobre-tarjeta).

Recordatorio importante: este validador mide las cajas DECLARADAS en
el XML del pptx, no el texto realmente renderizado. Puede decir
PASS aunque haya texto desbordándose visualmente de su caja (esto ya
pasó varias veces durante el desarrollo del motor). Para contenido con
texto de longitud variable, siempre complementar con revisión visual
convirtiendo a PDF/imagen y mirando el resultado.
