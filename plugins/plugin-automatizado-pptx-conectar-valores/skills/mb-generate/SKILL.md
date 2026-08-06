---
name: mb-generate
description: Genera presentaciones PowerPoint (.pptx) desde cero para Conectar Valores S.A.S., aplicando un grid derivado al estilo Müller-Brockmann (módulo/canaleta basados en el cuerpo de texto, zonas verticales estrictas) y la identidad visual exacta de CV (paleta navy/dorado, tipografía Cambria/Arial, footer corporativo, logo). Usa este skill cuando el usuario pida crear, generar o armar una presentación, deck, o slides nuevas para Conectar Valores — especialmente si menciona "grid", "formato CV", "estilo Conectar Valores", o si ya ha usado este skill antes en la conversación. NO usar para formatear o corregir un .pptx que ya existe (eso es /mb-format) — este skill es solo para construir contenido nuevo.
---

> **Entorno requerido: Claude Code (CLI)**
> Este skill ejecuta comandos Node.js en el sistema de archivos local. No
> funciona en claude.ai chat ni en ningún entorno sin acceso a Bash. Si
> intentas usarlo fuera de Claude Code CLI, Claude puede generar el código
> del script pero no puede ejecutarlo ni entregar el archivo .pptx.

# mb-generate — Generador de presentaciones CV con grid Müller-Brockmann

## Qué hace este skill

Construye archivos `.pptx` completos (portada + slides de contenido) que cumplen
estrictamente las reglas de marca de Conectar Valores S.A.S. y un sistema de grid
derivado matemáticamente del tamaño del cuerpo de texto (estilo Müller-Brockmann),
en vez de posicionar elementos a ojo.

Todo el motor vive en `engine/` y ya pasó un backtest de 15 casos límite
(ver `tests/stress_test_full.js`). **No reescribas la lógica de cálculo de
zonas o de overflow de texto a mano** — usa siempre las funciones de
`engine/mb_engine.js`, `engine/mb_cover.js`, y `engine/mb_text_fit.js`. Si
necesitas un layout que no existe todavía, créalo como una función nueva
en el motor (siguiendo el mismo patrón de protección contra desborde que
ya usan `addCard` y `layoutFlowchart`), no como código inline en el script
del deck.

## Flujo de trabajo

1. **Reunir el contenido con el usuario.** Antes de generar nada, confirma:
   título del documento/proyecto, número aproximado de slides, y qué tipo
   de contenido va en cada una (texto+tabla, tarjetas, flowchart, etc.).
   Si el usuario no especifica, pregunta — no asumas estructura.

2. **Resolver imágenes si la portada las necesita.** Usa
   `engine/mb_images.js` → `fetchHeroImage(query, outDir, opts)` para
   buscar y descargar una foto real de Pexels relacionada al tema. Nunca
   uses imágenes ilustrativas dibujadas a mano ni placeholders en el
   entregable final — solo está bien dejar el placeholder si el usuario
   explícitamente no quiere buscar una imagen.

3. **Inferir el nombre de proyecto para el footer.** Usa
   `engine/mb_engine.js` → `inferProjectName(title)` para derivar el texto
   corto del footer ("GRUPO EJEMPLO", "ACME", etc.) a partir del título
   de la primera slide de contenido — nunca lo hardcodees a mano salvo que
   el usuario pida un texto de footer distinto explícitamente.

4. **Construir slide por slide** usando los helpers del motor (ver
   `reference/engine_api.md` para la firma completa de cada función):
   - Portada → `mb_cover.js::buildCover()`
   - Cabecera de cualquier slide de contenido → `deriveGrid()` +
     `deriveHeaderBlock()` + `addHeader()`
   - Texto/bullets → `addText` normal de pptxgenjs dentro de
     `grid.cols.L`/`grid.cols.R`
   - Tablas → `tblCell()` + `addTable` de pptxgenjs
   - Tarjetas de color → `addCard()`
   - Diagramas de flujo verticales → `layoutFlowchart()` + `addFlowStep()`
   - Notas al pie → `addNotes()`
   - Footer → `addFooter()`

5. **Validar antes de entregar.** SIEMPRE correr
   `engine/mb_validator.js` contra el archivo generado:
   ```
   node engine/mb_validator.js /ruta/al/archivo.pptx
   ```
   Si reporta problemas, corregir antes de mostrar el archivo al usuario.
   Esto detecta overlaps geométricos y violaciones de zona, pero **no
   sustituye la revisión visual** — para cualquier layout con texto
   variable (tarjetas, flowchart, títulos largos), conviértelo a imagen
   (`soffice --headless --convert-to pdf` + `pdftoppm`) y mírala antes de
   entregar, especialmente si el contenido es inusualmente largo o corto.

6. **Manejar overflow honesto.** `layoutFlowchart()` puede devolver
   `overflow: true` si el contenido no cabe ni al tamaño de fuente mínimo.
   Cuando eso pase, NO fuerces que todo quepa comprimiendo a un tamaño
   ilegible — trunca a los pasos que sí caben y avísale al usuario, o
   sugiere dividir en dos slides. El patrón de manejo está en
   `tests/stress_test_full.js`, caso `C3_flowchart_9_pasos_overflow_esperado`.

## Reglas de marca no negociables

- Paleta: ver `PALETTE_CV` en `mb_engine.js` (navy `1A2744`, dorado `B8862A`,
  fondos de tabla `F2F4F7`/`E4E8EF`). No improvisar otros colores.
- Tipografía: Cambria para títulos/serif, Arial para todo lo demás.
- El footer SIEMPRE es la barra navy de 3 segmentos
  (confidencialidad / contexto / proyecto·página) — nunca solo el logo.
- El logo CV va arriba a la derecha, alineado verticalmente con el título,
  en slides de contenido — nunca en el footer (se ve demasiado pequeño ahí).
- La portada sigue SIEMPRE el template fijo de `buildCover()` — no es un
  layout que se rederive por slide, es identidad de marca fija.
- Ninguna zona (footer, notas, contenido) puede invadir a otra. Los
  respiros entre zonas (`GAP_FN`, `GAP_CN` en `mb_engine.js`) son
  intencionales, no recortarlos para ganar espacio.

## Configuración requerida

Este skill necesita una API key de Pexels en `config/api_keys.json` para
buscar imágenes reales. Ver `INSTALL.md` en la raíz del plugin para cómo
configurarla. Si el archivo no existe o la key es inválida,
`fetchHeroImage()` lanzará un error — en ese caso, avisa al usuario y
ofrece continuar con un placeholder o pedirle que suba su propia imagen.

## Bitácora de observaciones — aprendizaje sin auto-edición

Este plugin lo usa un equipo, no una sola persona, así que el motor
(`engine/*.js`) **nunca se modifica automáticamente** en base a lo que
pase durante una generación — eso requeriría que alguien apruebe cada
cambio en tiempo real, lo cual no escala con varias personas generando
decks de forma independiente.

En su lugar, existe `logs/observations.md`: una bitácora **local a esta
copia del plugin** donde se registran observaciones reales de uso (tanto
problemas como casos límite que funcionaron bien). Cada persona del
equipo que instale el plugin tendrá su propia copia de este archivo,
que crece con su propio uso — no se sincroniza automáticamente entre
máquinas. Si en algún momento se quiere consolidar lo observado por
varias personas, alguien tiene que compartir su `observations.md`
manualmente (por el canal que ya use el equipo) para revisarlo junto
con el resto.

**Cuándo escribir una entrada ahí** (ver criterios completos al inicio
de ese archivo):
- Si el validador dio ✓ PASS pero la revisión visual mostró algo
  desbordado, apretado, o mal alineado → registrar como PROBLEMA
- Si un caso de contenido real e inusual (título muy largo/corto, tabla
  con muchas/pocas filas, cuerpo de tarjeta extenso, etc.) se vio bien
  sin que tuvieras que intervenir manualmente → registrar como ÉXITO
- Generaciones rutinarias sin nada inusual no necesitan entrada

**Lo que NUNCA debe pasar:** modificar `engine/mb_engine.js`,
`mb_cover.js`, `mb_text_fit.js`, `mb_validator.js`, o `mb_images.js`
durante una generación normal de deck, incluso si detectas un problema
claro. Anótalo en la bitácora y sigue adelante con un workaround puntual
para ese deck si hace falta (ej. ajustar manualmente el texto de esa
slide específica) — la corrección del motor en sí es una decisión
separada que toma el dueño del plugin, no cada generación individual.

Si el usuario en la conversación actual es explícitamente el dueño del
plugin (lo dice, o están revisando la bitácora juntos a propósito), ahí
sí se puede seguir el protocolo normal de edición del motor: cambio →
`node tests/stress_test_full.js` (debe seguir en verde) → revisión
visual del caso afectado → listo.

## Limitaciones conocidas

- `fitTextLines()` (en `mb_text_fit.js`) es una heurística calibrada
  empíricamente (CHAR_WIDTH_RATIO 0.58), no una medición real de texto
  renderizado. Es confiable para los casos ya probados en el stress test,
  pero con fuentes o idiomas muy distintos al español/inglés podría
  necesitar recalibración — si ves desbordes nuevos, ese es el primer
  lugar a revisar.
- El motor no soporta todavía gráficos (charts) nativos de pptxgenjs con
  protección de overflow — si el usuario pide un gráfico, queda como
  trabajo manual sin las garantías de zona del resto del motor.
- `/mb-format` (auditoría de .pptx existentes) no está implementado
  todavía — ver carpeta `skills/mb-format/` para su estado.
