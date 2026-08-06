# IB Deck Playbook — Metodología Completa

Guía paso-a-paso para construir y extender slides estilo **MUFG-elegant** (banca de inversión) usando `python-pptx` + PowerPoint COM. **Genérica** — aplicable a cualquier proyecto (M&A, valoración, fairness opinion, DD report, pitch).

---

## 1. Filosofía del deck

- **Calidad 10/10 IB**: paleta sobria, tipografía Arial, márgenes generosos, tablas minimalistas, charts nativos sin clutter visual.
- **Densidad informativa**: cada slide responde una pregunta clara. Un comité de inversión debe poder leer cualquier slide en 30 segundos y extraer la conclusión.
- **Single-methodology consistency**: en valoración, el deck muestra UNA metodología principal con sensibilidades. NO comparativos confusos entre DDM/WACC/Múltiplos en deck content (eso va en anexos).
- **Datos auditables**: cada cifra proviene de una celda específica del modelo Excel oficial. Nunca inventar.
- **Verificación visual**: cada slide construido se exporta como PNG y se revisa antes de declarar done.

---

## 2. Style guide

### 2.1 Paleta (en `core/deck_style.py`)

```python
NAVY        = RGBColor(0x1F, 0x3D, 0x5C)   # primary titles, table headers
NAVY_DARK   = RGBColor(0x14, 0x2A, 0x40)
TEAL        = RGBColor(0x5D, 0x9E, 0xA7)   # accent, lines, sub-headers
TEAL_DARK   = RGBColor(0x3E, 0x7A, 0x82)
SAGE        = RGBColor(0xA8, 0xC5, 0xC0)
SAGE_BG     = RGBColor(0xE6, 0xEE, 0xEC)   # background sage tint
GRAY_LIGHT  = RGBColor(0xF4, 0xF6, 0xF7)   # alt rows, sidebar bg
GRAY_LINE   = RGBColor(0xD8, 0xDC, 0xDF)   # borders
GRAY_MID    = RGBColor(0xA9, 0xAE, 0xB5)
GRAY_TEXT   = RGBColor(0x6B, 0x71, 0x79)   # secondary text
DARK_TEXT   = RGBColor(0x2A, 0x2D, 0x33)   # body text
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CREAM       = RGBColor(0xF6, 0xE8, 0xCB)   # highlight rows (totals)
ORANGE      = RGBColor(0xE0, 0x8E, 0x3C)   # accent variation
GREEN       = RGBColor(0x6F, 0xA0, 0x6F)
GOLD        = RGBColor(0xC9, 0xA2, 0x27)
```

**Uso convencional:**
- NAVY: títulos, headers tabla, KPI big values
- TEAL: chips de sección, bordes superior cards, líneas accent
- GRAY_LIGHT: filas alternadas tabla, paneles bottom
- CREAM: fila total highlight
- ORANGE/GREEN/GOLD: distinguir series en charts (uso sobrio, máx 3 colores per chart)

### 2.2 Tipografía

- **Font**: `Arial` (NUNCA Calibri ni Times)
- **Body text**: 9-10pt
- **Subtitle (Rectangle 90)**: 14-16pt regular TEAL
- **Section titles (chip-style)**: 8.5pt bold NAVY uppercase con tracking
- **Sub-section titles**: 10.5-11pt bold NAVY con línea TEAL bajo
- **Table headers**: 9-10pt bold WHITE on NAVY
- **KPI big values**: 13-15pt bold NAVY
- **Footer**: 8pt GRAY_TEXT

### 2.3 Canvas y márgenes (inches)

```python
SLIDE_W = 13.333    # 16:9 widescreen
SLIDE_H = 7.5
LM = 0.65           # left margin
RM = 0.65           # right margin
TM = 1.15           # top (debajo del master title bar)
BM = 0.40           # bottom (encima del footer)
USABLE_W = 12.033   # SLIDE_W - LM - RM
USABLE_H = 5.95     # SLIDE_H - TM - BM
```

### 2.4 Footer

- Posición fija **y=7.20"**
- Todo contenido del slide termina antes de **y=7.18"**
- Texto: `"ESTRICTAMENTE PRIVADO Y CONFIDENCIAL · PREPARADO POR · {EMPRESA} · {PROYECTO} · {n}"`

---

## 3. Layout patterns

5 patterns que cubren ~90% de slides IB. Coordinadas en inches.

### Pattern A: Snapshot (KPI strip + content)
```
y=1.40-1.95: intro paragraph
y=2.05-3.00: 4 KPI cards (cw = USABLE_W/4 each, gap 0.10")
y=3.05-6.00: main content (table o chart)
y=6.10-7.10: bottom panel (callout o nota analítica)
```

### Pattern B: Two columns (text + table/chart)
```
Mid section: split en 2 columnas
col_w = (USABLE_W - 0.35) / 2
Left:  bullets / texto rico (add_rich)
Right: tabla MUFG o chart
Bottom: callout NAVY
```

### Pattern C: Single chart slide
```
y=1.45-2.10: intro paragraph
y=2.30: section title
y=2.65: chart (h=3.10-3.30)
y=6.10-7.15: bottom panel con análisis (3-4 bullets)
```

### Pattern D: Two charts side-by-side
```
chart_w = (USABLE_W - 0.30) / 2
y=2.40-5.30: ambos charts (h=2.85)
y=5.85-7.15: panel inferior con interpretación
```

### Pattern E: Mini-charts grid 2x2
```
chart_h = 1.95
chart_w = (USABLE_W - 0.30) / 2
Top row:    y=2.15-4.50
Bottom row: y=4.65-7.00
⚠️ NO exceder y=7.00 (footer at 7.20)
```

---

## 4. Helpers core (`core/deck_style.py`)

| Helper | Uso |
|---|---|
| `set_subtitle(slide, text)` | Reemplaza placeholder `xxxx` en Rectangle 90 con subtítulo |
| `add_textbox(slide, x, y, w, h, text, **kwargs)` | Caja de texto plano (size, bold, color, align, vanchor, line_spacing, font) |
| `add_rich(slide, x, y, w, h, runs, line_spacing)` | Texto multi-run con formato mixto. `runs = [[(text, dict_format), ...], ...]` (lista de párrafos, cada uno lista de tuplas) |
| `add_rect(slide, x, y, w, h, fill, line, line_w)` | Rectángulo (cards, panels, separators) |
| `add_filled_text(slide, x, y, w, h, text, fill, color, ...)` | Caja rellena con texto centrado (callouts, KPI big values) |
| `add_table_mufg(slide, x, y, w, h, data, header_size, body_size, col_aligns, first_col_bold, highlight_rows, total_rows)` | Tabla MUFG con headers NAVY, alt rows GRAY_LIGHT, opcional highlight + totals CREAM |
| `add_kpi_strip(slide, kpis, y)` | Strip horizontal de KPI cards (4 default) |
| `add_kpi_sidebar(slide, kpis, x, y, ...)` | Sidebar vertical de KPIs |
| `add_section_chip(slide, x, y, text)` | Chip "SECCIÓN N · NOMBRE" — TEAL uppercase tracking |
| `add_section_title(slide, x, y, w, text)` | Sub-section title con línea TEAL bajo |
| `add_horizontal_divider(slide, x, y, w)` | Línea horizontal teal |
| `add_arrow(slide, x, y, w, h, ...)` | Flecha → conector |
| `add_footer(slide, page_num)` | Footer estándar |

### Chart helpers (en `examples/{PROJECT}/scripts/`)

| Helper | Tipo de chart |
|---|---|
| `add_bar_chart(slide, l, t, w, h, categories, values, bar_color, value_format)` | COLUMN_CLUSTERED |
| `add_stacked_bar(slide, l, t, w, h, categories, series_data, colors)` | COLUMN_STACKED |
| `add_line_chart(slide, l, t, w, h, categories, values, line_color)` | LINE con data labels |
| `add_pie_chart(slide, l, t, w, h, categories, values, colors)` | DOUGHNUT |
| `add_tornado_row(slide, label, low_pct, high_pct, y, ...)` | Fila de tornado chart (drawn shapes) |
| `add_football_bar(slide, x_left, x_right, y, ..., low_pct, high_pct)` | Range bar de football field |

---

## 5. Workflow paso-a-paso

### A. Inspect deck state
```python
from pptx import Presentation
p = Presentation(SRC)
print('Total:', len(p.slides))
for i, s in enumerate(p.slides):
    titles = [sh.text_frame.text for sh in s.shapes
              if sh.has_text_frame and sh.text_frame.text.strip() not in ('', 'xxxx')]
    print(i, titles[:1])
```

### B. Build new section
1. **Pedir al usuario** que duplique los templates necesarios via PowerPoint UI (NO `python-pptx` deepcopy — rompe imágenes).
2. **Crear** `examples/{PROJECT}/scripts/proj_s{N}_slides.py` con un builder por slide:
```python
def slide_X_name(slide):
    set_subtitle(slide, "Título descriptivo")
    add_section_chip(slide, LM, 1.05, "Sección N · Título Sección")
    # intro textbox
    add_textbox(slide, LM, 1.45, USABLE_W, 0.65,
                "Intro explanation...", size=10, color=DARK_TEXT, line_spacing=1.32)
    # KPI strip / chart / table según pattern
    # bottom panel callout
    add_footer(slide, X)
```
3. **Crear** `proj_s{N}_main.py` con reset+build:
```python
import sys, shutil
sys.path.insert(0, r"C:\path\to\IB_Deck_Kit\core")
sys.path.insert(0, r"C:\path\to\IB_Deck_Kit\examples\{PROJECT}\scripts")
from pptx import Presentation
from deck_style import *
from proj_s_N_slides import slide_X, slide_Y, ...

shutil.copy(ONEDRIVE_DECK, LOCAL_DECK)
prs = Presentation(LOCAL_DECK)

BASE_NAMES = {"think-cell data - do not delete",
              "Marcador de número de diapositiva 1",
              "Rectangle 90", "object 10",
              "Straight Connector 2", "Imagen 1", "Picture 2"}

# Reset templates (preserve only BASE_NAMES)
for idx in range(START, END):
    s = prs.slides[idx]
    for sh in list(s.shapes):
        if sh.name not in BASE_NAMES:
            sh.element.getparent().remove(sh.element)
    # Reset Rectangle 90 text to "xxxx"
    for sh in s.shapes:
        if sh.name == "Rectangle 90" and sh.has_text_frame:
            p = sh.text_frame.paragraphs[0]
            if p.runs:
                p.runs[0].text = "xxxx"
                for r in p.runs[1:]: r.text = ""

# Run builders
builders = [(idx, slide_X), ...]
for idx, b in builders:
    b(prs.slides[idx])

prs.save(LOCAL_OUT)
shutil.copy(LOCAL_OUT, ONEDRIVE_DECK)
```
4. **Build**: `python proj_s{N}_main.py`
5. **Export PNG** via PowerShell COM:
```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open("DECK_PATH", $true, $false, $false)
$pres.Slides.Item($N).Export("OUT.png", "PNG", 1600, 900)
$pres.Close(); $ppt.Quit()
```
6. **Read PNG** con la herramienta Read y verifica visualmente
7. **Iterar** hasta 10/10

### C. Targeted fix de 1 slide
- Crear `proj_fix_sX.py` que solo resetea+rebuild ese índice
- Más rápido que rebuild completo de la sección

### D. Edición textual deck-wide (typos, sub-strings)
- Crear script tipo `proj_fix_text.py` con `replace_run_text()`
- Útil para fixes sistémicos (ej. "DOS DE CINCO" → numeros correctos)

```python
def replace_in_paragraph(para, find, replace):
    full = ''.join(r.text for r in para.runs)
    if find in full:
        para.runs[0].text = full.replace(find, replace)
        for r in para.runs[1:]: r.text = ''
        return True
    return False
```

### E. Audit numérico
1. Cargar modelo: `wb = openpyxl.load_workbook(MODEL, data_only=True)`
2. Cruzar cada cifra del deck contra celda específica
3. Documentar en `AUDIT_NOTES.md` con tabla GO + issues
4. Aplicar fixes para issues GRAVE/MENOR

---

## 6. Tricks importantes

### 6.1 Templates duplication
- **NO** usar `python-pptx` deepcopy — rompe relaciones de imágenes (logos, mapas)
- **SI** usar PowerShell COM: `$pres.Slides.Item(N).Duplicate()`
- O pedir al usuario duplicar via PowerPoint UI

### 6.2 Reset de slide preservando estructura base
```python
BASE_NAMES = {"think-cell data - do not delete",
              "Marcador de número de diapositiva 1",
              "Rectangle 90", "object 10",
              "Straight Connector 2", "Imagen 1", "Picture 2"}
keep_count = {n: 0 for n in BASE_NAMES}
to_remove = []
for sh in slide.shapes:
    if sh.name in BASE_NAMES:
        keep_count[sh.name] += 1
        if keep_count[sh.name] > 1:
            to_remove.append(sh.element)
    else:
        to_remove.append(sh.element)
for el in to_remove:
    el.getparent().remove(el)
```

### 6.3 OneDrive locks
```python
shutil.copy(ONEDRIVE_PATH, LOCAL_PATH)
prs = Presentation(LOCAL_PATH)
# ... edits ...
prs.save(LOCAL_OUT)
shutil.copy(LOCAL_OUT, ONEDRIVE_PATH)
```

### 6.4 Charts data labels overlap
- Si valores cercanos al baseline → labels colisionan
- Solución: deshabilitar data labels y enable `value_axis` con format
- O: posicionar labels en `INSIDE_BASE` en lugar de `OUTSIDE_END`

### 6.5 Bottom panel overflow footer
- Footer está fijo en y=7.20
- Todo contenido debe terminar antes de y=7.18
- Si overflow: shrink chart height por 0.30-0.50", subir bottom panel
- O remover una fila de tabla / un bullet

### 6.6 Pie chart % display
- Si format="0.00%", pasar valores como **decimales** (0.4492 no 44.92)
- Format `"0.00%"` multiplica por 100 internamente

### 6.7 Run text fragmentation
- Algunos runs vienen split (`"SECCION " | "UNO" | " DE CINCO"`)
- Para text replacement: combinar runs primero
```python
full = ''.join(r.text for r in para.runs)
if find in full:
    para.runs[0].text = full.replace(find, replace)
    for r in para.runs[1:]: r.text = ''
```

### 6.8 Chart fill color override
- Theme override puede pisar tu color
- Force `series.format.fill.solid()` y luego set RGB explicit
- Para línea sin marker: `series.format.line.fill.background()` después de set color

### 6.9 PowerPoint COM kill antes de operaciones
- Si hubo un error o un PowerPoint quedó abierto → archivo locked
- Antes de cada export PNG: `Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force`

### 6.10 PNG cache stale
- A veces PowerPoint COM exporta una versión cacheada
- Si el PNG no refleja los cambios: kill PowerPoint, esperar 1 sec, retry
- O exportar desde el archivo LOCAL (no OneDrive) para evitar sync delay

---

## 7. Errores comunes y fixes

| Error | Causa | Fix |
|---|---|---|
| `PackageNotFoundError` | OneDrive lock | `shutil.copy` a local antes de abrir |
| Slides en blanco tras rebuild | Indices desfasados | Verificar `len(prs.slides)` y ajustar `range()` |
| Charts no respetan paleta | Theme override | `series.format.fill.solid()` + RGB explicit |
| Pie muestra 4492% | Format multiplica | Pasar `0.4492` no `44.92` con format `"0.00%"` |
| Bottom table choca footer | Overflow | Shrink chart h, subir panel, recortar contenido |
| Subtitle 2 líneas | Texto largo | Acortar título, partir en `set_subtitle` y add_textbox |
| Data labels colisionan | Series cercanas | Disable labels, enable `value_axis` |
| Run text replace falla | Run fragmentado | Combinar runs primero (sec 6.7) |
| Imágenes no aparecen tras dup | python-pptx deepcopy | Usar PowerShell COM Duplicate() |
| Texto chino/ruso/raro en cells | Encoding | `python -c` con `sys.stdout.reconfigure(encoding='utf-8')` |

---

## 8. Estructura típica deck IB (templates por sección)

```
1. Cover
2. Disclaimer / Confidentiality
3. Table of Contents
4. Section divider 1: Resumen Ejecutivo / Executive Summary
5. Resumen Ejecutivo (1-3 slides)
6. Section divider 2: Perfil de la Compañía / Company Profile
7-N. Perfil slides (snapshot, localización, infraestructura, estructura societaria, hitos, ESG, permisos, pólizas, investment highlights)
N+1. Section divider 3: Mercado y Contrapartes / Market & Customers
   ... slides de mercado, offtakers, ratings
N+2. Section divider 4: Contratos / Contracts
   ... slides de contratos por contraparte
N+3. Section divider 5: Históricos Financieros / Historical Financials
   ... snapshot, ingresos, costos, EBITDA, util neta, deuda
N+4. Section divider 6: Valoración / Valuation
   ... resumen, supuestos, proyecciones, cuentas reserva, CFADS, deuda, pagos restringidos / dividendos, metodología, tasa descuento, resultados
N+5. (sub-sección) Sensibilidades — tornado, football field, escenarios
N+6. Contactos / Final
N+7. Appendix (opcional)
```

---

## 9. Checklist final

### Por slide
- [ ] PNG exportado y verificado visualmente
- [ ] Sin overflows footer
- [ ] Sin texto wrap en títulos
- [ ] Cifras auditadas vs modelo
- [ ] Paleta consistente
- [ ] Tipografía Arial

### Deck completo
- [ ] Cover con fecha correcta
- [ ] Section dividers numerados consistentemente (UNO/DOS/.../N DE N)
- [ ] TOC alineado con secciones
- [ ] Cifras headline (Equity Value, Deuda Neta, Ke, etc.) verificadas vs modelo
- [ ] AUDIT_NOTES.md con GO/NO-GO documentado
- [ ] Backup `deck_ref{N}.pptx` antes de cambios mayores
- [ ] Memory `project_*.md` actualizada
- [ ] Deploy a OneDrive verificado

---

## 10. Filosofía IB

> Un comité de inversión decide en 30 segundos por slide. La densidad informativa, la jerarquía visual, y la honestidad numérica son irrenunciables. Si dudas si un slide es 10/10, pregúntate: **¿imprimirías esto y lo enviarías al CIO de un fondo de pensiones?**

> No exageres. No infles. No inventes. Si no sabes, di que no sabes. La credibilidad de un deck IB se construye con cada cifra correcta — y se destruye con un solo número equivocado.

---

## 11. Estructura de un caso

Cada proyecto vive en `examples/{PROYECTO}/` con esta estructura:

- `DATA_REFERENCE.md` — todas las cifras canónicas con su celda fuente
- `AUDIT_NOTES.md` — audit GO/NO-GO + issues
- `scripts/` — builders (style + builders por sección + main + fixes)
- `deck/` — deck final `.pptx`
- `source_materials/` — PDFs, imágenes y material fuente
- `examples_png/` — screenshots de slides para verificación visual
- `memory/` — memorias persistentes del proyecto

Los casos reales no se versionan en este repositorio: contienen información
confidencial de clientes. Crea tu carpeta de proyecto siguiendo esta estructura.
