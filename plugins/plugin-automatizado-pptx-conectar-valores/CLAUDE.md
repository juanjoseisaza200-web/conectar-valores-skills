# CLAUDE.md — plugin-automatizado-pptx-conectar-valores

Contexto vivo de este proyecto. Léelo al abrir esta carpeta para entender
qué es, qué se ha hecho, y qué falta — sin tener que releer todo el
historial de conversación.

## Qué es esto

Plugin de Claude Code con dos skills para Conectar Valores S.A.S.
(Medellín, banca de inversión):

- **mb-generate** — genera presentaciones .pptx desde cero aplicando un
  grid derivado al estilo Müller-Brockmann (módulo/canaleta basados en el
  tamaño del cuerpo de texto) más la identidad visual exacta de CV (paleta
  navy/dorado, Cambria/Arial, footer corporativo de 3 segmentos, logo).
- **mb-format** — audita y reformatea un .pptx existente aplicando el grid MB y la identidad CV. IMPLEMENTADO — template "subsidiary_profile" (perfiles de empresa subsidiaria) probado 25/25 contra un deck de valoración real.
  que cumpla el mismo grid. No implementado todavía, solo existe como
  placeholder documentando el problema (ver skills/mb-format/SKILL.md).

## Estado: mb-generate

Completo y backtested. 15/15 casos límite pasan
(skills/mb-generate/tests/stress_test_full.js), cubriendo los 4 patrones
de layout: header+texto/tablas, tarjetas, flowchart, portada.

### Bugs reales encontrados y corregidos durante el desarrollo

Estos son importantes de recordar porque todos comparten la misma causa
raíz — vale la pena tenerlos en mente si se agregan layouts nuevos:

1. Título principal se desbordaba sobre el subtítulo — pptxgenjs no hace
   autofit de texto al generar el archivo; si el texto solicitado a un
   tamaño de fuente dado ocupa más líneas de las que la caja reservó,
   simplemente se desborda visualmente sin error. Fix: estimar líneas con
   fitTextLines() ANTES de reservar la altura, no asumir 1-2 líneas fijo.

2. mb_cover.js tenía su propia heurística de título (umbral de
   caracteres fijo length > 38) en vez de usar la misma lógica del motor
   general — quedaba desincronizada y era más débil. Fix: extraer
   fitTextLines() a mb_text_fit.js, usado por ambos.

3. El CHAR_WIDTH_RATIO inicial (0.52) subestimaba el ancho real de texto
   BOLD — un título que la heurística decía que cabía en 2 líneas en
   realidad ocupaba 3 en PowerPoint real. Recalibrado a 0.58 tras
   verificar contra 3 casos reales conocidos.

4. El piso de fuente (Math.max(12, ...)) podía exceder el tamaño
   SOLICITADO — pedías 11pt y a veces te devolvía 12pt (más grande, no
   más chico). Fix: el piso ahora es Math.min(6, fontSizeSolicitado),
   nunca puede superar lo que se pidió.

5. addCard() no protegía el CUERPO de la tarjeta contra desborde, solo
   el título. Un cuerpo muy largo se salía por debajo de la tarjeta de
   color hasta invadir el footer. Fix: mismo patrón de reducción de
   fuente que ya tenía el título, aplicado también al cuerpo.

6. layoutFlowchart() comprimía las cajas silenciosamente cuando el
   contenido no cabía ni al tamaño de fuente mínimo — el "último
   recurso" escalaba alturas hacia abajo SIN PISO, produciendo cajas
   más chicas que su propio texto (label y body se superponían dentro
   de la misma caja). Fix: eliminado el escalado silencioso. Ahora
   devuelve overflow true más overflowAmount, y el caller decide qué
   hacer (truncar pasos, avisar al usuario, dividir en 2 slides) —
   nunca se inventa espacio que no existe.

Patrón general de la lección: el validador geométrico
(mb_validator.js) NUNCA detectó ninguno de estos 6 bugs por sí solo,
porque mide cajas declaradas en el XML, no el texto realmente
renderizado por PowerPoint/LibreOffice. Los 6 se encontraron
exclusivamente por revisión visual de los casos límite del stress
test. Conclusión: cualquier layout nuevo necesita su propio caso de
stress test con texto deliberadamente largo/corto, renderizado a
imagen y revisado a ojo, antes de considerarse confiable — el
validador geométrico es un complemento, no un sustituto.

## Estado: mb-format

No empezado. Ver skills/mb-format/SKILL.md para el análisis de por qué
es un problema fundamentalmente distinto (parte de contenido ajeno y
arbitrario en vez de controlar la generación desde el origen) y qué se
necesitaría para construirlo.

Antes de empezar esto, hay una decisión de producto pendiente: si el
caso de uso es "corregir mis propios PPTX mal hechos" (acotado) vs
"corregir cualquier PPTX de un tercero" (mucho más abierto). Preguntarle
a JJ.

## Decisiones de diseño que vale la pena recordar

- pptxgenjs usa slides de 10in x 5.625in con LAYOUT_16x9 — NO 13.33in.
  Este fue el primer bug encontrado en el proyecto (contenido
  desbordándose a la derecha) y vale la pena no repetirlo si se empieza
  un layout nuevo desde cero sin usar deriveGrid().
- El logo CV vive como PNG con fondo transparente
  (skills/mb-generate/assets/logo_cv.png, originalmente
  Untitled_design.png con fondo negro removido vía PIL). Es cuadrado
  (2000x2000), aspect ratio 1:1.
- Las imágenes de portada vienen de Pexels vía API (key gratuita en
  config/api_keys.json, NUNCA hardcodeada en el código del motor). El
  plan es que esta key se comparta directamente entre el equipo de JJ
  sin pasar por GitHub — si eso cambia, hay que mover la key fuera del
  versionado.
- El footer SIEMPRE es la barra navy de 3 segmentos
  (confidencialidad/contexto/proyecto-página) — un diseño anterior con
  solo el logo en la esquina fue rechazado explícitamente por el
  usuario.
- El nombre de proyecto en el footer se infiere del título de la
  primera slide de contenido (inferProjectName()), no se hardcodea —
  ej. "Valoración del Grupo Ejemplo" → "GRUPO EJEMPLO".

## Archivos de referencia rápida

- skills/mb-generate/SKILL.md — punto de entrada, cuándo y cómo usar
  el skill
- skills/mb-generate/reference/engine_api.md — firma completa de cada
  función del motor, para no tener que releer el código fuente
- skills/mb-generate/tests/stress_test_full.js — los 15 casos límite,
  útil como ejemplo de uso real de cada helper del motor
- INSTALL.md (raíz del plugin) — instrucciones de instalación para
  alguien nuevo en el equipo
