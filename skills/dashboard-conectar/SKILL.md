---
name: dashboard-conectar
description: Construye dashboards financieros ejecutivos en Excel (.xlsm) con el estandar visual de Conectar Valores S.A.S. — estilo PEL/one-pager, paleta corporativa, charts dinamicos con filtro de periodo, sin VBA. Usar SIEMPRE que se pida crear, rediseñar o ampliar un dashboard/tablero sobre un modelo financiero. Procedimiento por fases con gates: ESCANEA el modelo hoja por hoja, RAZONA a profundidad que metricas/ratios/flujos presentar (con su celda fuente), DISEÑA el layout, CONSTRUYE con la libreria COM probada (tools/dash_lib.py), agrega interactividad y AUDITA (KPIs identicos al modelo, ESF=0, colores corporativos, sin sparklines, sin solapes). Cero alucinacion: toda cifra trazable a una celda.
---

# Dashboard financiero — estándar Conectar (estilo PEL)

**Un dashboard es PURA PRESENTACIÓN: NO modifica el modelo.** Lee fórmulas vivas del modelo y las muestra; los KPIs del modelo deben quedar **idénticos** antes y después. Diseñada para ejecutarse sin errores: sigue las fases EN ORDEN, usa `tools/dash_lib.py` (código probado en producción — no reescribas COM a mano), respeta la paleta y las lecciones de `reference/`, y NO declares terminado nada sin PASS de `tools/audit_dashboard.py`.

## Prohibiciones absolutas (memorizar antes de empezar)
1. **NUNCA `taskkill /f /im EXCEL.EXE`** ni matar Excel — cierra los libros abiertos del usuario (ya pasó). SIEMPRE `cv.abrir(ruta, kill=False)` (instancia aislada DispatchEx) + `xl.Quit()` en `finally`. `RPC_E_CALL_REJECTED` = sleep+reintento, no matar.
2. **NUNCA modificar el modelo.** El dashboard solo añade/edita una hoja de presentación + una hoja oculta de helpers. Los KPIs (EBITDA, utilidad, caja, deuda, TIR, ESF) deben salir **idénticos** en la auditoría. Si cambian → rompiste algo.
3. **NUNCA sparklines** (el usuario los detesta). El auditor falla si encuentra uno.
4. **NUNCA leer una celda/check tras calcular sin `cv.calc_wait`** (CalculationState==0).
5. **NUNCA asumir posiciones de fila** — búscalas por ETIQUETA (`dl.fB`/`dl.find_colE`). Las filas se corren entre versiones del modelo y la hoja se **compacta sola** en el primer guardado (ver `reference/charts-com.md` §row-collapse).
6. **NUNCA dejar una serie en color default del tema.** Re-apuntar `Series.Values` RESETEA el color → tras CUALQUIER repunte, `dl.apply_colors`. Ver `reference/interactividad.md`.
7. **NUNCA form-control DropDowns vía COM** (`.Value`/`.ListIndex` lanzan error; `RemoveAllItems` corrompe la LinkedCell). Interactividad = **validación de datos EN CELDA** + rangos OFFSET. Ver `reference/interactividad.md`.
8. **NUNCA charts modernos vía `ChartType=`** (treemap 117/waterfall 119/sunburst/mapa). Usa clásicos; waterfall = bridge apilado. Ver `reference/charts-com.md`.
9. **NUNCA inventar una métrica sin celda fuente.** Toda cifra del dashboard se ancla a una celda del modelo verificada en FASE 0 (cero alucinación). Si el modelo no la calcula, o se construye con fórmula trazable, o se declara como NO disponible.
10. **NUNCA sobrescribir** la base ni el dashboard vigente — copia nueva `... vN.xlsm`.

## Setup (una vez por script)
```python
import sys, os
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\dashboard-conectar\tools"))
import dash_lib as dl          # importa cv_model como dl.cv
```
Ejecutar con `$env:PYTHONIOENCODING="utf-8"; python script.py`. Leer también el `CLAUDE.md` del proyecto y `aprender-de-errores/LECCIONES.md`.

## FASE 0 — ESCANEAR el modelo completo (sin tocar nada)
1. `python tools/scan_dashboard.py "ruta_modelo.xlsm"` → mapa hoja por hoja: lista de hojas, etiquetas de fila con sus valores (F escalar, Q=2026, AD=2039) de las hojas clave (EEFF, Deuda, WK, Waterfall, Ingresos, OPEX, CAPEX, Inputs), selector de escenario (`Inputs_C!F6`), fila del Check ESF, unidades.
2. **Unidades**: el modelo suele estar en COP **miles**; el dashboard reporta COP **mil MM** = modelo/1e6. Confírmalo (un KPI conocido).
3. Construir el **mapa de fuentes**: por cada cifra candidata, su `Hoja!celda`. Verificar 2-3 anclas conocidas (EBITDA, utilidad neta, ESF).
4. **GATE 0**: tienes el mapa `métrica → Hoja!celda` de TODO lo relevante, las unidades confirmadas, y sabes dónde está el motor de escenarios y el Check ESF. Si no → seguir escaneando.

## FASE 1 — RAZONAR el contenido (el corazón de la skill)
Con el mapa de FASE 0, **razonar A PROFUNDIDAD qué presentar** — no copiar un layout fijo, sino pensar qué le importa a quien lee (lender / comité / inversionista) según el tipo de transacción. Guía obligatoria: **`reference/razonar-contenido.md`** (universo de métricas por dimensión + reglas de selección).
1. Identificar: tipo de deal, riesgos clave, audiencia. Recorrer las 10 dimensiones (ejecutivo, P&L, retornos, deuda/cobertura/covenants, apalancamiento, liquidez, working capital, operativo, calidad de cartera, flujo de caja) y para cada una decidir qué métricas/ratios/charts entran, **cada uno con su celda fuente**.
2. Reglas: cada número trazable; sin redundancia (no repetir el mismo dato en dos charts — si ya está arriba, abajo va OTRO dato relevante); preferir charts que cuenten una historia (tendencia/desapalancamiento) sobre comparaciones planas; señalar lo que el modelo NO permite (ej. LLCR≡PLCR) en vez de fabricar.
3. Escribir el **PLAN DE CONTENIDO** (`templates/plan-contenido.md`): secciones → para cada una sus KPIs/filas (con `Hoja!celda`), charts (tipo, series con fuente), y la banda KPI ejecutiva de 8 tarjetas.
4. **GATE 1**: presentar el plan al usuario con las fuentes; aprobación explícita antes de construir.

## FASE 2 — DISEÑAR el layout (PEL one-pager)
Estructura canónica (de arriba a abajo), detalle en `reference/paleta-y-formato.md` §layout:
`FILTRO DE PERIODO` → banda **KEY INFORMATION** (8 tarjetas KPI a todo el ancho, número grande + sub▲▼) → **gráficas hero** (3 charts clave: cobertura/DSCR, paydown, covenant) → **PANELES TEMÁTICOS** (cada uno = banda navy + **tabla compacta C:P** a la izquierda [años 2026-2039 en cols C:P, etiqueta en col B] + **gráfica R:AD** a la derecha que lee la tabla compacta) → **secciones de detalle** con charts hero → **cascada de caja** al final.

## FASE 3 — CONSTRUIR (script COM con dash_lib, re-ejecutable)
1. Copia versionada del modelo → será el dashboard. Crear hoja `Dashboard` (o editar la existente) + hoja oculta `_DashViz` (helpers de charts no-FAST).
2. Tablas: `dl.compact_mirror(...)` (C:P = espejo de `Hoja!Q:AD`/1e6). KPIs: `dl.kpi_band(...)`. Bandas: `dl.band(...)`. Charts: `dl.make_chart(...)` (clásicos, leen `$C$r:$P$r`, X=`$Q$7:$AD$7`). Cascada: `dl.waterfall_bridge(...)`. Anclar todos: `dl.anchor_all_charts` (Placement=xlMove).
3. Paleta y tipos difíciles: `reference/paleta-y-formato.md` + `reference/charts-com.md`. **Sin sparklines.**
4. Si algo falla a mitad: NO parchear — corregir el script y re-ejecutar desde la copia limpia.

## FASE 4 — INTERACTIVIDAD (filtro de periodo, opcional pero recomendado)
1. `dl.add_period_filter(wb, sh, hv)` — crea 2 dropdowns de validación en celda (Desde/Hasta año), nombres OFFSET dinámicos, y re-apunta TODAS las series de serie-temporal. Sin VBA.
2. **Tras el repunte: `dl.apply_colors(sh, ref)`** con `ref=dl.capture_colors` tomado ANTES (o desde la versión previa). Verificar con `tools/audit_dashboard.py --diff-colors base.xlsm`.
3. Escenarios: los charts leen fórmulas vivas → cambiar `Inputs_C!F6` + recerrar el escenario actualiza todo solo (no requiere código). Detalle: `reference/interactividad.md`.

## FASE 5 — CERRAR Y AUDITAR (gates duros, no negociable)
1. `cv.calc_wait(xl, full=True)` → `cv.circulares(wb)` == `[]`.
2. `cv.cerrar_modelo(xl, wb)` → checks=0 (0 iteraciones = pura presentación, correcto).
3. **GATE 5**: `python tools/audit_dashboard.py "dashboard vN.xlsm" --base "modelo.xlsm"` → PASS de:
   KPIs idénticos al modelo base · ESF=0 en 21 cols · 0 errores · 0 circulares · 0 sparklines · 0 charts en colores default del tema · 0 solapes de charts · filtro en rango completo por defecto. FAIL → corregir script, re-Fase 3, volver aquí.
4. Render visual (`dl.render_range` / grab por rango) de cada sección para revisar layout/colores con los ojos. Captura de charts flotantes: `Range.CopyPicture(1,2)` (no AddChart2 headless).

## FASE 6 — DOCUMENTAR
Memo .md de qué muestra cada sección + su fuente; actualizar memoria del proyecto (versión vigente, posiciones, decisiones); registrar cualquier lección nueva en `aprender-de-errores/LECCIONES.md`.

## Referencia rápida
- `reference/razonar-contenido.md` — universo de métricas por dimensión + reglas para decidir QUÉ presentar (FASE 1).
- `reference/paleta-y-formato.md` — paleta BGR, tipografía, tarjetas KPI, tabla compacta, bandas, layout PEL.
- `reference/charts-com.md` — todas las lecciones COM de charts (tipos difíciles, waterfall, donut, row-collapse, `&` en nombre de hoja).
- `reference/interactividad.md` — filtro de periodo sin VBA (validación en celda + OFFSET), reset de color al repuntar, dropdowns prohibidos.
- `tools/dash_lib.py` — librería de construcción (bandas, tarjetas, tablas compactas, charts, cascada, filtro, colores, render).
- `tools/scan_dashboard.py` — escáner del modelo (FASE 0).
- `tools/audit_dashboard.py` — auditor PASS/FAIL (FASE 5).
- `templates/plan-contenido.md` — plantilla del plan de contenido (FASE 1).
- `tools/cv_model.py` — librería COM base (abrir/calc_wait/cerrar_modelo/circulares/errores/check_esf). **Incluida en la skill** (copia) para portabilidad 100%: funciona sola; si existe `modelaje-conectar` usa esa como fallback. Detalle de cierre/escenarios en `modelaje-conectar/reference/com-y-escenarios.md` (si está disponible).
