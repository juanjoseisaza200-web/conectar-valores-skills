# Interactividad sin VBA (filtro de periodo + escenarios)

Objetivo: gráficos dinámicos que respondan a un filtro de periodo, **sin VBA ni Python** → abre y funciona en cualquier Excel sin habilitar permisos. Lo hace `dl.add_period_filter`.

## Mecánica (validación en celda + rangos OFFSET)
1. **Controles = validación de DATOS en celda** (NO form controls): `C9`=Desde año, `G9`=Hasta año, con `Validation.Add(Type=3, Formula1="=Dashboard!$Q$7:$AD$7")`. La celda muestra el año elegido (sin los bugs de los form controls).
2. **Celdas índice** en `_DashViz`: `A50=MATCH(C9, años, 0)`, `A51=MATCH(G9, años, 0)`, `A52=START(=A50)`, `A53=CNT(=MAX(1, A51-A50+1))`.
3. **Nombres dinámicos OFFSET** (scope libro):
   - `f_x = OFFSET(Dashboard!$Q$7, 0, START-1, 1, CNT)` (eje años, recortado).
   - `f_<r> = OFFSET(Dashboard!$C$<r>, 0, START-1, 1, CNT)` por cada fila compacta charteada; `g_<r>` para filas de `_DashViz`.
4. **Re-apuntar series**: cada serie de tiempo (Values era `Dashboard!$C$r:$P$r`, X `$Q$7:$AD$7`) → `s.Values="='file'!f_r"`, `s.XValues="='file'!f_x"`. Series por categoría (TIR/NETGO) y cascada NO se tocan.
5. Al mover Desde/Hasta, MATCH→índices→OFFSET→los charts se recortan a la ventana. Default = rango completo (Desde=1er año, Hasta=último).

## ⚠️ TRAMPAS (todas verificadas en producción)
- **Re-apuntar `Series.Values` RESETEA el color al tema.** Por eso, tras CUALQUIER repunte → `dl.apply_colors(sh, ref)`, con `ref` capturado ANTES (`dl.capture_colors`) o leído de la versión previa buena. Verificar con `audit_dashboard.py --diff-colors`.
- **NUNCA `Series.Formula="=SERIES(...)"`** para repuntar — reconstruye la serie y pierde TODO el formato. Usar solo los setters `.Values` y `.XValues`.
- **Form-control DropDowns (`ws.DropDowns`) son una trampa vía COM**: `.Value`/`.ListIndex` lanzan `Unable to set...`; el display no sincroniza con la LinkedCell; `RemoveAllItems` ESCRIBE 0 en la LinkedCell (rompe el default a 1 año). → usar SIEMPRE validación en celda.
- Cambiar la ventana, `calc_wait` y `cerrar_modelo` **NO** resetean colores (verificado) — el único que resetea es el repunte de Values. Así que basta colorear una vez DESPUÉS del último repunte.
- Tras tocar dropdowns/series, re-verificar colores (`--diff-colors` o `dl.capture_colors` comparado).

## Escenarios (vía F6, sin código)
Los charts leen fórmulas vivas del modelo (a través de las tablas compactas C:P que espejan `Hoja!Q:AD`). Por eso, cuando el usuario cambia el escenario en `Inputs_C!F6` y re-cierra ese escenario (`cv.cerrar_modelo`), **todo el dashboard se actualiza solo** — KPIs y charts. No requiere botón ni macro. Si se quisiera un botón "recalcular escenario" o un what-if con recálculo en vivo, ESO sí requeriría VBA (y habilitar confianza en el proyecto VBA) — proponerlo aparte solo si el usuario lo pide.

## Por qué no Python-in-Excel / VBA
- Python-in-Excel (=PY): potente pero depende de M365 con la función habilitada y cálculo en la nube; riesgo de compatibilidad con .xlsm pesado de macros.
- VBA: inyectar código vía COM exige habilitar "Confiar en el acceso al modelo de objetos de proyectos VBA"; correrlo exige macros habilitadas.
- El enfoque nativo (validación + OFFSET) no necesita nada de eso → es el default. Reservar VBA/Python solo para what-if con recálculo en vivo si el usuario lo exige.
