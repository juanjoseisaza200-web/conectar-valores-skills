# Conectar Valores — Skills & Kits

Colección de skills, plugins y kits para Claude usados en Conectar Valores S.A.S.
(banca de inversión, project finance, modelaje financiero y documentos corporativos).

## Contenido

### `skills/` — Skills de Claude Code / Claude Desktop

| Skill | Qué hace |
|---|---|
| `conectar-docx-creator` | Crea documentos Word nuevos desde cero con el formato corporativo VF3 (Arial, paleta Navy/Dorado/Mint, márgenes Letter, tablas y footer CV). |
| `conectar-docx-format-applier` | Aplica el formato VF3 a contenido existente (texto, markdown, JSON o un `.docx` con otro formato). No genera contenido. |
| `modelaje-conectar` | Modelaje financiero en Excel con el estándar CV (metodología FAST). Fases con gates, librería COM (`tools/cv_model.py`) y auditor automático. |
| `auditoria-md-conectar` | Auditoría de modelos financieros con mentalidad de MD: protocolo de 17 fases, backtesting PxQ, CFADS/DSCR/covenants, sensibilidades. |
| `dashboard-conectar` | Dashboards ejecutivos en Excel (`.xlsm`) estilo one-pager, charts dinámicos sin VBA, con auditoría de cifras contra el modelo. |
| `aprender-de-errores` | Mejora continua: registra lecciones cuando una verificación falla o el usuario corrige un error, y las consulta antes de repetir un tipo de tarea. |

### `plugins/` — Plugins de Claude Code

| Plugin | Skills incluidas |
|---|---|
| `plugin-automatizado-pptx-conectar-valores` | `mb-generate`, `mb-format`, `mb-generate-chat`, `mb-format-chat` — generación y reformateo de presentaciones `.pptx` con grid Müller-Brockmann e identidad visual CV. |

### `kits/` — Kits de trabajo (sin `SKILL.md`)

| Kit | Qué hace |
|---|---|
| `IB_Deck_Kit` | Toolkit para decks de banca de inversión con `python-pptx` + PowerPoint COM: style guide, playbook y librería `core/deck_style.py`. |
| `KIT_SENSIBILIDADES_CONECTAR` | Corre matrices de escenarios de sensibilidad sobre modelos Excel con macro de cierre de circulares. |
| `Kit_Notas_De_Voz` | Transcripción de notas de voz a texto. |

## Instalación

**Skills en Claude Code** — copiar (o enlazar) la carpeta de la skill a `~/.claude/skills/`:

```bash
ln -s "$PWD/skills/modelaje-conectar" ~/.claude/skills/modelaje-conectar
```

**Plugin pptx** — agregar este repo como marketplace local:

```bash
claude plugin marketplace add ./plugins
```

**Dependencias de Python:**

```bash
pip install -r requirements.txt
```

Las skills de Excel y PowerPoint usan automatización COM, por lo que requieren Windows con Office instalado. Las de Word (`docx-creator`, `docx-format-applier`) funcionan en cualquier plataforma.

## Nota sobre material de clientes

Este repositorio contiene **solo las skills, su código y sus plantillas**. Todo el material
de clientes se mantiene deliberadamente fuera:

- Los casos de ejemplo (modelos `.xlsm`, decks de valoración, informes y material fuente)
  no se versionan.
- Los nombres de clientes, cifras de valoración y rutas locales fueron reemplazados por
  placeholders genéricos en la documentación y el código.
- El `.gitignore` bloquea `source_materials/`, `outputs/` y archivos `.xlsx` / `.xlsm`
  para que no entren por accidente.

Las plantillas `.docx` incluidas están vacías: solo contienen el estilo, la portada y el
footer corporativo, sin contenido de ningún informe.

Si vas a usar estas skills con un mandato real, mantén el modelo y los entregables fuera
del repositorio.
