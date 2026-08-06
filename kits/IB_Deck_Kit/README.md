# IB Deck Kit — Investment Banking Decks con Python

Toolkit completo para construir **decks de calidad banca de inversión** en formato MUFG-elegant usando `python-pptx` + PowerPoint COM. Reproducible, portable, transferible entre PCs.

## 🎯 Para qué sirve

- Construir decks de valoración / pitch / fairness opinion / DD report en cualquier proyecto
- Style guide MUFG: paleta navy + teal sobria, Arial, márgenes generosos, tablas minimalistas, charts nativos
- Pipeline reproducible: scripts Python que generan slides con datos auditables del modelo
- Transferible a otra sesión / otro PC sin perder contexto

## 📂 Estructura del kit

```
IB_Deck_Kit/
├── README.md                ← este archivo
├── CLAUDE.md                ← instrucciones para Claude (rol + reglas + workflow)
├── PLAYBOOK.md              ← metodología completa (style guide, layouts, workflow)
├── core/                    ← helpers genéricos REUSABLES en cualquier proyecto
│   └── deck_style.py        ← paleta, helpers (set_subtitle, add_textbox, add_table_mufg,
│                               add_kpi_strip, add_section_chip, add_section_title,
│                               add_filled_text, add_arrow, add_footer, add_rect, add_rich)
└── examples/
    └── {PROYECTO}/          ← una carpeta por proyecto (no versionadas)
        ├── DATA_REFERENCE.md
        ├── AUDIT_NOTES.md
        ├── scripts/         ← builders específicos del proyecto
        ├── deck/            ← deck final (.pptx)
        ├── source_materials/← PDFs, imágenes, mapas
        ├── examples_png/    ← screenshots
        └── memory/          ← memorias persistentes del proyecto
```

> Las carpetas de `examples/` contienen información confidencial de clientes y por eso
> no se versionan. El kit publica solo `core/` y la metodología.

## 🚀 Quick start (proyecto nuevo)

1. **Crea carpeta del proyecto nuevo**, ejemplo:
   `Desktop/IB_Deck_Kit/examples/MY_PROJECT/`

2. **Crea la estructura** descrita en `PLAYBOOK.md` §11 (scripts, deck, source_materials, etc.)

3. **Adapta los scripts**:
   - Edita los paths en `*_main.py` (SRC, DEST, sys.path)
   - Reescribe los builders con datos del nuevo proyecto
   - Mantén `core/deck_style.py` intacto (helpers genéricos)

4. **Abre Claude Code** en `IB_Deck_Kit/` y dile:
   > "Lee CLAUDE.md y PLAYBOOK.md. El proyecto es {nombre} en `examples/MY_PROJECT/`. Estoy listo para construir."

5. Listo.

## 📋 Quick start (continuar un proyecto)

1. Abre Claude Code en `IB_Deck_Kit/`
2. Dile: "Lee CLAUDE.md, PLAYBOOK.md, examples/{PROYECTO}/DATA_REFERENCE.md. Continúo en `examples/{PROYECTO}/`."

## ⚙️ Requisitos en el PC nuevo

```bash
python -m pip install python-pptx pymupdf openpyxl
```

- Microsoft PowerPoint instalado (necesario para COM: export PNG + duplicación slides)
- Bash (Git for Windows) + PowerShell

## 🎨 Style guide TL;DR

- **Paleta**: NAVY `#1F3D5C` · TEAL `#5D9EA7` · TEAL_DARK `#3E7A82` · GREEN `#6FA06F` · ORANGE `#E08E3C` · GOLD `#C9A227` · grays
- **Font**: Arial — body 9-10pt, headers 10.5-11pt bold, KPI big 13-15pt bold
- **Slide canvas**: 13.333" × 7.5" (16:9 widescreen)
- **Margins**: LM=RM=0.65", TM=1.15", BM=0.40"
- **Footer**: y=7.20" (todo contenido debe terminar antes de y=7.18")

Detalles completos en `PLAYBOOK.md`.

## 🆘 Soporte

- Errores comunes: `PLAYBOOK.md → Errores comunes y fixes`
- Estilo / paleta: `PLAYBOOK.md → Style guide`
- Layouts: `PLAYBOOK.md → Layout patterns`
- Helpers API: `core/deck_style.py` (docstrings inline)
- Caso completo: `examples/PEL_case/`
