---
name: conectar-docx-creator
description: Crea documentos Word NUEVOS desde cero con el formato exacto del Informe Peers Aseo VF3 de Conectar Valores S.A.S. Skill GENERADORA — recibe un brief (qué tipo de documento, qué secciones, qué tablas, qué contenido) y produce un .docx terminado con tipografía Arial, paleta hex CV (Navy/Dorado/Mint/Hairline), márgenes Letter 1.2/0.79", tablas con header texto Navy bold sobre blanco + underline dorado, filas con hairline gris, footer corporativo, todo verificado contra el XML del VF3 real. Soporta cualquier tipo de documento corporativo: análisis sectorial, informe de DD, memo, propuesta, factsheet, anexo técnico, reporte ejecutivo. Output queda en la carpeta que el usuario indique. Úsala cuando el usuario diga "genérame un documento sobre X", "crea un informe nuevo de Y", "construye un memo corporativo de Z" y quiera el formato CV.
---

# conectar-docx-creator

Skill **GENERADORA**: produce documentos Word **desde cero** con el formato visual literal del VF3 de Conectar Valores.

⚠️ **Diferencia clave vs `conectar-docx-format-applier`:**
- `conectar-docx-format-applier` toma un `.docx` existente y le aplica el formato (no toca contenido)
- **Esta skill** parte de un brief / outline y **crea contenido nuevo** con el formato VF3

## Comportamiento obligatorio del flujo

### Paso 1 — Conversar para definir el documento

**SIEMPRE preguntar antes de generar.** Recolectar:

1. **Tipo de documento** — ¿análisis sectorial, memo, informe DD, propuesta, factsheet, anexo, otro?
2. **Tema y alcance** — ¿de qué trata? ¿qué empresa/sector/proyecto?
3. **Audiencia** — ¿fondo internacional, banco, comprador, uso interno?
4. **Idioma** — español por defecto; inglés si el destinatario lo requiere
5. **Estructura deseada** — ¿secciones específicas? ¿el usuario aporta outline o pide que la skill proponga uno típico?
6. **Contenido por sección** — texto crudo, bullets, cifras, narrativa que el usuario quiera incluir
7. **Tablas** — para cada una: headers, rows (datos), si hay fila a destacar, nota al pie
8. **Carpeta destino** — ¿dónde guardar el `.docx`? Si no se indica, sugerir `C:\Users\Usuario\Desktop\` y confirmar
9. **Nombre del archivo** — formato sugerido: `YYYYMMDD_<Tema>_<Cliente>_VF.docx`

Usar la tool `AskUserQuestion` para preguntas con opciones (tipos de documento, idioma, audiencia).

### Paso 2 — Proponer outline antes de escribir

Antes de tocar el `.docx`, mostrar al usuario el outline propuesto en markdown:

```
# Outline propuesto

## Portada
- Logo + Título + Cliente + Fecha + Confidencial

## 1. Resumen Ejecutivo
- 5 bullets clave
- Párrafo cierre

## 2. Sección X
- Subsección 2.1
- Tabla 1: <headers>
...
```

Esperar OK del usuario o ajustes.

### Paso 3 — Recolectar contenido por bloque

Para cada sección/tabla del outline aprobado, preguntar (o pedir al usuario que pegue) el contenido crudo. Si el usuario delega ("rellénalo tú con plausible"), avisar explícitamente que las cifras serán plausibles pero NO reales.

### Paso 4 — Generar el .docx

Usar `templates/build_from_outline.py` que recibe un dict estructurado y produce el doc:

```python
import sys
sys.path.insert(0, r"C:\Users\Usuario\.claude\skills\conectar-docx-creator\templates")
from build_from_outline import build_document

doc_spec = {
    "output_path": "C:\\Users\\Usuario\\Desktop\\20260521_Analisis_X_VF.docx",
    "portada": {
        "empresa": "INTERASEO S.A.S E.S.P.",
        "titulo": "ANÁLISIS DE X",
        "subtitulo": "Comparativo Y vs Z",
        "descripcion": "...",
        "fecha": "Mayo 2026",
        "confidencial": True,
    },
    "secciones": [
        {"type": "h1", "text": "1. Resumen Ejecutivo"},
        {"type": "paragraph", "text": "..."},
        {"type": "bullets", "items": ["bullet 1", "bullet 2", ...]},
        {"type": "h2", "text": "1.1 ..."},
        {"type": "caption", "text": "Tabla 1 — ..."},
        {"type": "table", "headers": [...], "rows": [...], "highlight_row": N,
                          "col_widths_in": [...], "footnote": "Fuente: ..."},
        ...
    ]
}

build_document(doc_spec)
```

### Paso 5 — Auditar 8/8 antes de entregar

```bash
python critic/audit_against_vf3.py "<ruta_output>"
```

Reportar:
- Si pasa 8/8: confirmar ruta + ofrecer abrir
- Si falla: ajustar specs y regenerar

## Tipos de documento soportados

| Tipo | Estructura típica sugerida |
|---|---|
| **Análisis sectorial / Peers** | Portada → Glosario → Nota Metodológica → Resumen Ejec → Estructura Mercado → Mapeo Jugadores → Análisis Tarifario → Posicionamiento → Volúmenes → Infraestructura → Eficiencia → Subsidios → Análisis Financiero → Comparativo → Conclusiones → Referencias |
| **Memo corporativo (1-3pp)** | Portada simple → Asunto → Contexto → Análisis → Recomendación → Anexos |
| **Informe DD** | Portada → Resumen Ejec → Perfil del activo → Aspectos técnicos → Aspectos legales/regulatorios → Aspectos financieros → Riesgos → Conclusiones |
| **Propuesta / Pitch** | Portada → Resumen → Oportunidad → Equipo → Metodología → Cronograma → Pricing → Términos |
| **Factsheet (1-2pp)** | Portada compacta → Highlights → Métricas clave → Estructura financiera → Términos |
| **Anexo técnico** | Portada → Contexto → Tablas/Charts → Notas metodológicas → Fuentes |

## Specs heredadas del VF3 (idénticas a conectar-docx-format-applier)

- **Fuente:** Arial únicamente
- **Tamaños autorizados:** sz=17 (8.5pt), 18 (9pt — cuerpo), 19 (9.5pt), 20 (10pt), 22 (11pt), 24 (12pt), 26 (13pt), 32 (16pt portada)
- **Paleta hex:** Navy `#17375E`, Dorado `#C9A449`, Hairline `#BFBFBF`, Mint `#E7F0EF` (única autorizada)
- **Página:** Letter, márgenes 1.20" / 0.79" / 0.79" / 0.79"
- **Tabla VF3:** header blanco con texto Navy bold 9.5pt + bottom dorado; filas blancas con texto negro 9pt + bottom hairline; fila destacada Mint con texto Navy bold

Ver `reference/vf3_dna.md` para detalle XML literal.

## Estructura de la skill

```
conectar-docx-creator/
├── SKILL.md                    ← este archivo
├── reference/
│   └── vf3_dna.md              ← ADN visual literal (copia del format-applier)
├── templates/
│   ├── vf3_styles.py           ← Librería helpers (compartida con format-applier)
│   └── build_from_outline.py   ← Constructor genérico desde dict
├── assets/
│   └── vf3_blank_template.docx ← VF3 con cuerpo vaciado (168 estilos)
├── critic/
│   └── audit_against_vf3.py    ← Auditor 8 checks
├── examples/
│   └── outline_sectoral.json   ← Ejemplo de spec para análisis sectorial
└── (sin iterations/ — output va a la carpeta que el usuario indique)
```

## Restricciones absolutas

- **SIEMPRE preguntar antes de generar** — outline, contenido, carpeta destino
- **SIEMPRE mostrar outline en markdown para aprobar** antes de tocar Word
- **NUNCA inventar cifras críticas** sin avisar al usuario
- **NUNCA usar la skill informe-peers-ib** (deprecada por specs erradas)
- **SIEMPRE auditar 8/8** antes de entregar
- **Carpeta destino debe ser confirmada por el usuario** — sugerir Desktop si no indica
- **Cumplir tono CV** si es informe sectorial/peers: ver `reference/vf3_dna.md` + tone rules CV (no tono crediticio, sin frases prohibidas)
