---
name: peers-format-applier
description: Aplica el formato visual EXACTO del Informe Peers Aseo VF3 (Conectar Valores S.A.S.) a cualquier contenido. NO es una skill generadora — es una skill FORMATEADORA. Recibe contenido bruto (texto plano, markdown, JSON estructurado o un .docx con formato distinto) y devuelve un .docx con tipografía Arial, paleta hex CV exacta, márgenes Letter 1.2/0.79", tablas con header texto Navy bold sobre blanco + underline dorado, filas con hairline gris, fila destacada Mint, footer corporativo, todo verificado contra el XML del VF3 real. Úsala cuando el usuario quiera "ponerme esto en formato VF3", "aplica el formato del análisis de peers a este documento", "convierte a formato Conectar Valores", "dale formato CV a este contenido".
---

# peers-format-applier

Skill **formateadora** que aplica el formato visual literal del **Informe Peers Aseo VF3** (`20260513_Informe_Peers_Aseo_VF3.docx`) a cualquier contenido nuevo.

⚠️ **Diferencia clave vs `informe-peers-ib`:**
- `informe-peers-ib` es **generadora de contenido** con formato (estructura 15 secciones, narrativa peers)
- **Esta skill (`peers-format-applier`)** es **solo aplicadora de formato** — no genera contenido, solo aplica el "skin" VF3 al contenido que recibe.

## Cuándo invocar

- "Aplica el formato del análisis de peers a este texto"
- "Ponme esto en formato VF3 / Conectar Valores"
- "Dale formato CV a este documento"
- "Convierte este Word a formato peers"
- "Formatea esta tabla como las del análisis de peers"

## Inputs aceptados

| Input | Cómo entregarlo | Procesamiento |
|---|---|---|
| Texto plano / markdown | string directo o ruta a `.md`/`.txt` | Parsea headings (`#`), bullets, tablas markdown |
| JSON estructurado | dict con `sections: [...]` y `tables: [...]` | Mapeo directo a helpers |
| `.docx` con otro formato | ruta a `.docx` | Extrae texto + tablas + headings, re-formatea |
| Tabla suelta (CSV / list of lists) | ruta o estructura | Aplica `add_table_vf3` |

## Cómo usar

### Paso 1 — Cargar especificación del formato (obligatorio)
Antes de tocar contenido, leer:
1. `reference/vf3_dna.md` — ADN visual literal (paleta, tamaños, tablas, márgenes)
2. `assets/vf3_dna/*.json` — datos crudos por si hay ambigüedad

### Paso 2 — Abrir template clonado
```python
from templates.vf3_styles import open_blank_template, add_h1, add_h2, add_paragraph, add_table_vf3, add_caption, add_footnote
doc = open_blank_template()  # ← hereda los 168 estilos del VF3, footer, theme, márgenes
```

### Paso 3 — Inyectar contenido
Solo usar los helpers de `vf3_styles.py` — NO crear estilos paralelos:
- `add_h1(doc, "...")` → Heading 1 (12pt bold Navy centrado)
- `add_h2(doc, "...")` → Heading 2 (10pt bold Navy izquierda)
- `add_h3(doc, "...")` → Heading 3 (9.5pt bold Navy izquierda)
- `add_paragraph(doc, "...")` → cuerpo Arial 9pt justificado
- `add_bullet(doc, "...")` → bullet Arial 9pt justificado
- `add_caption(doc, "Tabla N — ...")` → caption Navy bold 8.5pt centrado
- `add_table_vf3(doc, headers=[...], rows=[...], highlight_row=N)` → tabla VF3 exacta
- `add_footnote(doc, "Fuente: ...")` → nota al pie Slate italic 8.5pt

### Paso 4 — Validar con el crítico
Antes de entregar, correr el subagente con `critic/critic_prompt.md`:
- Compara XML del output vs ADN VF3 (mismas fuentes, sizes, colores, bordes)
- Score sobre items críticos (paleta + bordes + fonts)

## Specs literales que la skill respeta

### Página
- Letter (8.5×11"), márgenes 1.20" / 0.79" / 0.79" / 0.79"
- Footer: *"Informe de Peers • [Sector] | Confidencial • Uso interno Conectar Valores S.A.S."*

### Tipografía
- Arial únicamente (Verificado en VF3: 100% Arial)
- 9pt cuerpo (dominante 1,252 runs en VF3)
- 9.5pt headers tabla / H3
- 8.5pt captions / footnotes
- 12pt H1, 10pt H2

### Paleta (única autorizada)
- Navy `#17375E` — títulos, headers tabla, fila destacada (texto)
- Dorado `#C9A449` — únicamente borde inferior header tabla (sz=12 = 1.5pt)
- Hairline `#BFBFBF` — únicamente borde inferior filas cuerpo (sz=2 = 0.25pt)
- Mint `#E7F0EF` — único fill autorizado (fila destacada cliente)

### Prohibidos (NO usar):
- `#EDEDED` bandas alternadas
- `#4472C4` azul Office
- `#70AD47` verde Office
- `#FF0000`, `#E67E22`, `#27AE60` (skill anterior los tenía mal)

### Tabla VF3 — ADN literal
- Header: **blanco con texto Navy bold 9.5pt centrado** + borde inferior dorado 1.5pt
- Cuerpo: **blanco con texto negro 9pt regular centrado** + borde inferior hairline 0.25pt
- Fila destacada: **fill Mint con texto Navy bold 9pt centrado**
- Tabla centrada en página, layout fijo, sin bordes laterales ni externos
- **NO hay bandas alternadas**
- **NO hay fila Total con fondo Navy**

## Modos de uso

### Modo `format-doc`: docx → docx
```python
# Usuario aporta un .docx con otro formato
input_doc = "ruta/al/input.docx"
output_doc = format_existing_docx(input_doc, output_path="formatted.docx")
```

### Modo `format-from-markdown`: md → docx
```python
markdown_text = """
# Sección 1
Párrafo de prueba...

## 1.1 Sub-sección
- Bullet uno
- Bullet dos

| Peer | Ingresos | EBITDA |
|------|----------|--------|
| Empresa A | 850 | 32% |
| **Cliente** | 920 | 34% |
| Empresa B | 720 | 28% |
"""
doc = format_markdown_to_vf3(markdown_text, highlight_keyword="Cliente")
```

### Modo `format-table-only`: lista → tabla VF3 en docx
```python
table_dict = {
    "headers": ["Peer", "Ingresos", "EBITDA"],
    "rows": [...],
    "highlight_row": 1,  # 0-based
    "col_widths_in": [2.0, 1.5, 1.5],
    "caption": "Tabla 1 — Universo de peers",
    "footnote": "Fuente: ..."
}
```

## Validación obligatoria post-generación

Correr el script `critic/audit_against_vf3.py` que verifica:
1. **Fonts:** solo Arial
2. **Sizes:** 9pt dominante (>50% de runs)
3. **Text colors:** solo `17375E` (Navy) + heredado negro
4. **Fills:** solo `E7F0EF` (Mint) en celdas destacadas — 0 EDEDED / 0 colores ajenos
5. **Border colors:** solo `BFBFBF` (hairline) + `C9A449` (dorado header)
6. **Border sizes:** 2 (hairline 0.25pt) + 12 (dorado 1.5pt) — sin sz=4
7. **Alignment:** center dominante
8. **Margins:** exactos 1728/1138/1138/1138 twips

**PASA** si los 8 checks dan OK contra ADN VF3.

## Restricciones absolutas

- **NUNCA crear estilos nuevos** — usar solo los heredados del template clonado
- **NUNCA pintar header de tabla en Navy** — VF3 lo tiene blanco con texto Navy
- **NUNCA agregar bandas alternadas EDEDED** — VF3 no las tiene
- **NUNCA forzar Justify en celdas de tabla** — VF3 usa center
- **NUNCA usar tamaños fuera de** {17, 18, 19, 20, 24} half-points (8.5, 9, 9.5, 10, 12 pt)
