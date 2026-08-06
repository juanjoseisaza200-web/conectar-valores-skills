# Protocolo Excel COM y escenarios — detalle operativo

Usar SIEMPRE la librería `tools/cv_model.py` — no reescribir este código a mano.

## Sesión COM
```python
import sys; sys.path.insert(0, r"<skill>\tools")
import cv_model as cv
xl, wb, W = cv.abrir(ruta)          # mata EXCEL previo, DispatchEx, calc manual
...ediciones...
cv.calc_wait(xl, full=True)          # SIEMPRE esperar CalculationState==0 antes de leer
```
- `DispatchEx` NO se registra en el ROT → no intentar re-adherirse con GetActiveObject/GetObject en un proceso posterior. Si el script muere, matar Excel y re-ejecutar end-to-end.
- Patrón de script: **re-ejecutable desde copia limpia** (copiar base → aplicar TODO → verificar → guardar versión nueva). Jamás estados a medias.
- `.Formula` siempre en sintaxis EN-US. PasteSpecial: fórmulas −4123, formatos −4122, valores −4163.
- Macros del cliente con MsgBox (p. ej. `CopyBloque`): NO ejecutarlas con Application.Run (bloquean); usar `cv.cerrar_modelo(xl, wb)` que replica su lógica.
- Circulares: `cv.circulares(wb)` tras cualquier cambio de fórmulas (con DisplayAlerts=False Excel NO avisa). Los `$A$1` con `CELL("filename")` son benignos y ya están filtrados.
- Wrapper de stdout: NO usar `sys.stdout = io.TextIOWrapper(...)` en scripts que importan librerías que imprimen; ejecutar con `python -X utf8`.

## Cierre del modelo (rompe-circularidad + sizing)
La hoja `Macros` tiene filas `_COPY` (fórmulas vivas: FCF, GMF, IRA, saldos de intereses) y `_PASTE` (valores pegados). El cierre = iterar copy→paste-valores hasta que `Check_Macros` y `Check_TotalMacros` = 0, intercalando el bloque `Repayment` (sizing de deuda). Está implementado en `cv.cerrar_modelo(xl, wb)`; nombres requeridos: `Scenario, Variables, Macros_Copy, Macros_Paste, Repayment_Copy, Repayment_Paste, Check_Macros, Check_TotalMacros, Check_Repayment, Run_sizing`.

**Cuándo re-cerrar (obligatorio)**: tras cualquier cambio que toque caja, deuda, GMF, intereses o capital de trabajo; tras cambiar de escenario; tras cambiar inputs de un escenario. Si los resultados del cambio son matemáticamente idénticos (pura re-arquitectura), el cierre converge en 0 iteraciones — eso es la CONFIRMACIÓN, no un error.

## Escenarios
- Selector: el usuario escribe el NOMBRE en `Inputs_C!F6`; `I5 = MATCH(F6, K5:X5, 0)` resuelve la columna. Nombres en `Inputs_C!K5:X5` (ej.: Caso Base / −100bps / −200bps / +WHT Intereses).
- Todos los inputs: `F = INDEX(K:X, $I$5)` → F6 conmuta todo el modelo.
- Macros refleja: `F46 ('Scenario') = Inputs_C!I5`; almacena UN bloque de cierre POR escenario ("EJ 1/2/3…", offsets `2×esc + Variables×(esc−1)`; Repayment `2×esc + 10×(esc−1)`).
- **Cerrar un escenario NO cierra los demás.** Para presentar varios: por cada escenario → poner F6 → `cv.cerrar_modelo` → auditar → siguiente.
- Inputs nuevos: llenar K:X COMPLETO (mismo valor en todas las columnas si no varía por escenario); fórmulas de calibración con columnas ABSOLUTAS antes de propagar K:X.

## Inserción de filas (cirugía)
1. `cv.assert_lbl(hoja, fila, "texto ancla")` ANTES y DESPUÉS de cada insert.
2. Insertar de abajo hacia arriba por hoja (`sh.Rows("92:139").Insert()`).
3. Fórmulas/nombres existentes se auto-ajustan; las fórmulas NUEVAS se escriben con posiciones FINALES (offset +n para toda fila ≥ punto de inserción, acumulando por insert).
4. Insertar DENTRO de la sección temática correspondiente — nunca al final de la hoja.
5. No borrar ni reordenar filas existentes.

## Verificación y entrega
```
python tools/auditar_modelo.py "ruta.xlsm" --filas "WK:93-134,398-426;TX:204-227" [--cerrar]
```
PASS obligatorio antes de declarar terminado. Además: snapshot de KPIs antes/después (utilidad, patrimonio, caja, deuda, TIR, EBITDA) y explicar toda variación; guardar como versión nueva `... vN.xlsm`.
