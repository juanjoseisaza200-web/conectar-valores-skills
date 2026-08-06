---
name: modelaje-conectar
description: Skill de modelaje financiero en Excel con el estandar de Conectar Valores S.A.S. (metodologia FAST). Usar SIEMPRE que haya que crear, modificar, auditar o cerrar un modelo financiero .xlsx/.xlsm de Conectar o de sus clientes - antes de tocar una sola celda. Procedimiento por fases con gates obligatorios, libreria COM probada (tools/cv_model.py), auditor automatico (tools/auditar_modelo.py), plantillas literales de formulas, convenciones de formato y mecanica de escenarios/cierre por macro.
---

# Modelaje financiero — estándar Conectar (FAST)

**Diseñada para ejecutarse sin errores por cualquier modelo.** No improvises: sigue las fases EN ORDEN, usa las herramientas de `tools/` (código probado — no reescribas COM a mano), copia las plantillas de `reference/formulas.md` literalmente, y NO declares terminado nada sin PASS del auditor.

## Prohibiciones absolutas (memorizar antes de empezar)

1. **NUNCA** proyectar en EEFF/balance — ahí solo `=IF(flag; Motor!fila; Hist!fila)`.
2. **NUNCA** constantes sueltas en hojas — escalares en `Inputs_C` (K:X completo + `F=INDEX(K:X,$I$5)`), series en `Inputs_Years`; la hoja importa con `=Inputs_C!$F$x`.
3. **NUNCA** una fila con fórmulas distintas entre J y AD (principio FAST) ni fórmulas que empiecen en Q.
4. **NUNCA** `SUM($Q$x:Qx)` ni `COLUMNS($Q$x:…)` (se invierten al propagar) ni INDEX que pueda resolver columnas futuras (crea circular). Plantillas correctas: `reference/formulas.md` §4-5.
5. **NUNCA** leer una celda/check tras calcular sin `cv.calc_wait` (CalculationState==0).
6. **NUNCA** sobrescribir la base ni la versión vigente — siempre copia nueva `... vN.xlsm`.
7. **NUNCA** asumir: los flags se verifican por hoja leyendo etiquetas (en WK proyección=fila 7 y primera proyección=fila 8; en EEFF/TX/CAPEX proyección=fila 8); las cifras de un revisor se contrastan contra la versión vigente; la composición de un dato se verifica en la fuente por tercero.
8. **NUNCA** entregar sin: auditor PASS + snapshot KPIs antes/después explicado + cierre re-ejecutado.
9. Con clientes: **NUNCA** cambiar el modelo sin aprobación explícita del usuario (proponer con impacto cuantificado primero).

## Setup (una vez por script)
```python
import sys, os
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools"))
import cv_model as cv
```
Ejecutar con `python -X utf8` (sin re-envolver stdout). Leer también `LECCIONES.md` de la skill `aprender-de-errores` (sección Excel) y el `CLAUDE.md` del proyecto.

## FASE 0 — Mapear (sin tocar nada)
1. Extraer del modelo: hojas, etiquetas de fila, fórmulas y valores de las zonas a tocar (TSV o lecturas COM read-only), filas de flags POR HOJA, nombres definidos (`wb.Names`), y referencias cruzadas a las celdas que se van a modificar (quién las consume).
2. Identificar: fila del Check ESF, mecánica de cierre (hoja Macros), selector de escenarios. Detalle: `reference/com-y-escenarios.md`.
3. **GATE 0**: puedes escribir de memoria el flag correcto de cada hoja implicada, las anclas con etiqueta de cada zona, y qué hojas consumen lo que vas a tocar. Si no → seguir mapeando.

## FASE 1 — Diseñar y aprobar
1. Por cada cambio: sección destino (dentro del tema, nunca al final de la hoja), filas a insertar, inputs nuevos (→ Inputs_C/Inputs_Years), fórmulas (elegidas de `reference/formulas.md`), impacto esperado en KPIs.
2. Presentar al usuario el diseño con impactos cuantificados. **GATE 1**: aprobación explícita (si es modelo de cliente).

## FASE 2 — Implementar (script end-to-end re-ejecutable)
Estructura del script: copiar base → asserts de anclas → inserts (de abajo hacia arriba; posiciones finales para fórmulas nuevas) → inputs en Inputs_C/Years → filas nuevas (formato por plantilla de rol + `cv.fast_fill`) → recableos → normalizar fuentes. Procedimientos exactos: `reference/com-y-escenarios.md` (inserción/COM) y `reference/formatos.md` (filas/colores).
Si algo falla a mitad: NO parchear el archivo — corregir el script y re-ejecutar desde la copia limpia.

## FASE 3 — Cerrar y auditar (gates duros)
1. `cv.calc_wait(xl, full=True)` → `cv.circulares(wb)` debe ser `[]`.
2. `cv.cerrar_modelo(xl, wb)` → checks = 0 (0 iteraciones = cambio de pura forma, correcto).
3. **GATE 3 (obligatorio, no negociable)**:
```
python tools/auditar_modelo.py "ruta vN.xlsm" --filas "WK:93-134,398-426;TX:204-227;..."
```
   PASS de los 7+ ítems (ESF=0 en 21 años, checks macro=0, 0 errores, 0 circulares, FAST uniforme, observados=0, sin rojos). FAIL → corregir el script, re-ejecutar Fase 2, volver aquí.
4. Snapshot KPIs antes/después (utilidad, patrimonio, caja, deuda, TIR, EBITDA): explicar TODA variación; cambio de pura forma ⇒ resultados IDÉNTICOS.

## FASE 4 — Escenarios (si aplica)
Cada escenario guarda su PROPIO cierre. Si el cambio toca inputs o si se van a presentar varios escenarios: por cada uno → poner el nombre en `Inputs_C!F6` → `cv.cerrar_modelo` → auditor. Mecánica completa: `reference/com-y-escenarios.md`.

## FASE 5 — Documentar y registrar
1. Memo .md + (si el usuario lo pide) Word formato VF3 con `conectar-docx-creator` auditado 10/10: cambios celda por celda, decisiones, impactos, pendientes.
2. Actualizar memoria del proyecto (posiciones nuevas, decisiones).
3. Si hubo cualquier corrección del usuario o verificación fallida → registrarla YA en `aprender-de-errores/LECCIONES.md`.

## Referencia rápida
- `reference/formulas.md` — plantillas literales (✔/✘) de TODAS las fórmulas permitidas.
- `reference/formatos.md` — colores, layout de fila, procedimiento de fila nueva, Word VF3.
- `reference/com-y-escenarios.md` — COM, cierre CopyBloque, escenarios, inserción de filas.
- `tools/cv_model.py` — librería (abrir, calc_wait, cerrar_modelo, fast_fill, auditorías parciales).
- `tools/auditar_modelo.py` — auditor PASS/FAIL del checklist completo.
- Errores históricos y reglas: `aprender-de-errores/LECCIONES.md`.
