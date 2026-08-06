# dashboard-conectar — instalación (para compañeros)

Skill para construir **dashboards financieros ejecutivos en Excel** (estilo one-pager Conectar Valores) sobre cualquier modelo .xlsm. Escanea el modelo, razona qué presentar, lo construye con la paleta y reglas corporativas, agrega filtro de periodo dinámico (sin VBA) y se autoaudita.

## Instalar (1 minuto)
1. Descomprimir este zip dentro de tu carpeta de skills de Claude Code:
   `C:\Users\<TU_USUARIO>\.claude\skills\`
   (debe quedar `...\.claude\skills\dashboard-conectar\SKILL.md`)
2. Listo. Claude la detecta sola. Invócala diciendo *"construye un dashboard sobre el modelo X"* o con `/dashboard-conectar`.

## Requisitos
- Windows + Excel de escritorio instalado.
- Python 3 con `pywin32` y `Pillow` (PIL):  `pip install pywin32 pillow`
- Es **autocontenida**: incluye su propia `cv_model.py` (librería COM). No necesita otras skills.

## Qué hace (fases)
ESCANEAR el modelo hoja por hoja → RAZONAR a profundidad qué métricas/ratios/flujos presentar (cada uno con su celda fuente, cero invención) → DISEÑAR layout → CONSTRUIR con `tools/dash_lib.py` → INTERACTIVIDAD (filtro de periodo) → AUDITAR (`tools/audit_dashboard.py`: KPIs idénticos al modelo, ESF=0, sin sparklines, sin solapes, colores corporativos).

## Prueba rápida
```
cd dashboard-conectar\tools
python scan_dashboard.py "ruta\a\tu_modelo.xlsm"          # mapa de fuentes
python test_skill.py "ruta\a\tu_modelo.xlsm"               # autotest de la librería (PASS)
```

Detalle completo en `SKILL.md` y `reference/`.
