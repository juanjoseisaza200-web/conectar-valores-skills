---
name: mb-format
description: Audita y reformatea un archivo .pptx EXISTENTE aplicando el grid Müller-Brockmann y la identidad de marca de Conectar Valores S.A.S. Usa este skill cuando el usuario suba o mencione un .pptx que ya existe y quiera aplicarle el formato CV, cambiar colores, reemplazar logos, o ajustar el grid. Template implementado: "subsidiary_profile" (slides de perfil financiero por empresa subsidiaria, como las slides 95-119 del deck de Interaseo). Invocar con /mb-format.
---

> **Entorno requerido: Claude Code (CLI)**
> Este skill ejecuta comandos Node.js y Python en el sistema de archivos local.
> No funciona en claude.ai chat ni en ningún entorno sin acceso a Bash. Si
> intentas usarlo fuera de Claude Code CLI, Claude puede generar el código
> pero no puede ejecutarlo ni entregar el archivo .pptx reformateado.

# mb-format — Auditoría y reformateo de PPTX existentes

## Estado: IMPLEMENTADO (template inicial: subsidiary_profile)

El skill funciona de punta a punta. Probado contra 25 slides reales del
deck de valoración de Grupo Interaseo (slides 95-119), con 25/25 detectadas
y reformateadas correctamente.

## Flujo de trabajo

1. El usuario sube (o indica la ruta de) un .pptx existente
2. Claude corre `node engine/mb_format.js <input.pptx> <rango> <output.pptx>`
   - `<rango>` puede ser: un número (`95`), un rango (`95-119`), o `all`
3. El motor detecta automáticamente el template de cada slide y aplica la
   transformación correspondiente. Slides sin template reconocido quedan
   con un placeholder de diagnóstico (no se pierden datos)
4. Se valida visualmente el output antes de entregarlo al usuario
5. Si algo se ve mal, se anota en `../mb-generate/logs/observations.md`

## Arquitectura (3 capas)

**Extractor** (`engine/extract_slide.py`) — lee el XML del .pptx original
con python-pptx y devuelve JSON estructurado con roles clasificados
(title, subtitle, company, description, fin_table, kpi_group, chart,
closing_text, section_header, footer_ignore, unknown). Nunca modifica
el archivo original.

**Reconstructor** (`engine/mb_reconstructor.js`) — toma el JSON del
extractor, detecta el template de la slide, y usa los helpers de
`../mb-generate/engine/mb_engine.js` para construir la slide nueva
con identidad CV aplicada. Lee el tamaño de slide real del archivo
(no asume 10x5.625 — puede ser 13.33x7.5 u otro) y pasa `slideW`/`slideH`
a `deriveGrid()` para que el grid se adapte al tamaño real.

**CLI** (`engine/mb_format.js`) — orquesta extractor + reconstructor,
maneja rangos de slides, y genera el .pptx de salida.

## Templates implementados

### subsidiary_profile
Slides de perfil financiero por empresa subsidiaria. Estructura:
- Col. izquierda: nombre empresa + descripción + sección Valoración
  con chart waterfall (como placeholder) + texto de cierre
- Col. derecha: sección Información financiera + 2 tablas lado a lado
  + tarjetas KPI (estilo outline minimalista, fondo claro + borde dorado)
  + sección Aspectos relevantes + bullets

Señales de detección: tiene `company` + `description` + `fin_table`
+ `closing_text`. El chart es opcional (algunas subsidiarias no tienen).

## Lo que NO hace todavía (y por qué)

- **Recolorar charts nativos**: el waterfall chart queda como placeholder.
  Recolorar un chart nativo de PowerPoint requiere editar el XML interno
  del chart (no las shapes de la slide), lo cual es una capa adicional
  de complejidad fuera del scope actual.
- **Templates no implementados**: slides de portada, índices, slides de
  texto plano, slides densas con logos múltiples. Para estas, mb-format
  devuelve un placeholder de diagnóstico con los roles detectados, para
  que el usuario sepa qué tiene y decida si vale la pena agregar ese
  template.
- **Slides con alta fragmentación de texto** (fragRatio > 1.5): texto
  cortado en muchos runs separados (ej. slide 4 "Índice" o slide 18 del
  deck de Interaseo) — el extractor las marca correctamente pero el
  reconstructor no tiene un template para "índice manual fragmentado".

## Cómo agregar un template nuevo

1. Inspeccionar 2-3 slides representativas del nuevo tipo:
   ```
   python3 engine/extract_slide.py <archivo.pptx> <idx> | python3 -m json.tool
   ```
2. Identificar los roles que aparecen consistentemente y que lo
   diferencian de los templates ya existentes
3. Agregar la condición en `detectTemplate()` en `mb_reconstructor.js`
4. Implementar la función `reconstruct<NombreTemplate>()` siguiendo el
   mismo patrón de `reconstructSubsidiaryProfile()`
5. Agregar un caso de prueba en `tests/` (equivalente al stress test de
   mb-generate pero con slides reales como input)

## Bitácora de observaciones

Comparte la bitácora con mb-generate: `../mb-generate/logs/observations.md`.
mb_format.js escribe automáticamente una entrada ahí cuando alguna slide
no tiene template reconocido (para no perder esa señal). Entradas manuales
siguen el mismo formato — etiquetar con [mb-format] para distinguirlas.
