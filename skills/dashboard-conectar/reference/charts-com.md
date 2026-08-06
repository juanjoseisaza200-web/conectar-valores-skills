# Lecciones COM de charts (las que más cuesta re-aprender)

## Protocolo COM (siempre)
- `cv.abrir(ruta, kill=False)` = `DispatchEx` (instancia AISLADA) → NUNCA toca los Excel del usuario. `xl.Quit()` en `finally`. **NUNCA `taskkill EXCEL`.**
- Tras `Calculate`/`CalculateFullRebuild`: esperar `cv.calc_wait(xl)` (CalculationState==0) antes de leer (lecturas rancias dan datos viejos).
- `RPC_E_CALL_REJECTED` / `'Ocurrió una excepción'`: sleep(0.5-1)+reintento, no abortar.
- Fechas a Excel: `Value2 = serial`, no `datetime`.
- **Hoja con `&` en el nombre** (ej. "CAPEX & PPE"): rompe el dispatch de `Range.PasteSpecial` en win32com → usar asignación directa `sh.Range("J{r}:AD{r}").Formula = jform` (Excel autoajusta relativas).

## Tipos de chart
- Clásicos por número: `xlColumnClustered=51`, `xlColumnStacked=52`, `xlLine=4`, `xlAreaStacked=76`, `xlBarClustered=57`.
- **Tipos modernos (treemap 117, waterfall 119, sunburst, mapa xlRegionMap 140) NO se crean con `ChartType=`.** Requieren `Shapes.AddChart2(-1, tipo, ...)` PERO `AddChart2 + SetSourceData(hoja oculta)` FALLA. → usar clásicos.
- **Mapa geográfico (Colombia, etc.) NO es automatizable** (Excel lo bloquea) → barra ordenada o treemap-como-barras de sustituto.
- **Waterfall fiable = BRIDGE APILADO clásico**: serie base invisible + subtotales NAVY + variaciones GOLD. Helper en `_DashViz` con columnas (paso, base, sub, var). `dl.waterfall_bridge` lo arma. Cierra exacto si base/var se calculan de los totales reales.

## Fuente de datos de las series
- **Series de tiempo leen la tabla compacta**: `Values="=Dashboard!$C$r:$P$r"`, `XValues="=Dashboard!$Q$7:$AD$7"` (X de la fila de años global, Y de la compacta C:P; ambas 14 puntos).
- Para mini-tablas por categoría (no años) o cascada: `Chart.SetSourceData(_DashViz!rango)` con categorías en la 1ª columna.
- Ocultar los datos originales Q:AD: fuente blanca + col Q angosta (ancho ~2.5).

## Donut / pie
- Color **por PUNTO**, no por serie: `ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB=...`.
- Etiqueta %: formatear las **celdas fuente** a `0.0%` + `DataLabels.NumberFormatLinked=True` (poner `NumberFormat` directo en DataLabels NO pega).

## Eje
- `Axes(2).TickLabels.NumberFormat="#,##0"` para valores; `"0.0%"` solo si la serie es %. **Al reemplazar un chart, el eje hereda el formato del anterior** → si era TIR (%), la nueva gráfica de saldos mostrará "45000%". Siempre fijar el formato del eje explícitamente tras reemplazar.
- Serie en eje secundario (línea sobre barras): `s.ChartType=xlLine; s.AxisGroup=2`.

## Estilo (dl.style_chart)
- ChartArea sin borde, fondo blanco, Arial 9 DGRAY. Título 11 bold NAVY. Leyenda abajo (Position=-4107) tam 8. Gridlines color GRID. PlotArea sin relleno.

## ★ Colapso de filas (row-collapse) — sorprende siempre
Tras construir secciones con gaps de filas vacías reservadas y hacer el **PRIMER guardado/cierre**, Excel **compacta las filas vacías UNA vez** → todas las bandas suben (decenas de filas). Los charts (posición absoluta) quedan donde estaban y el contenido se compacta hacia ellos → suele MEJORAR (quita whitespace). Es estable después (no se repite). Defensa:
- Buscar filas SIEMPRE por etiqueta (`dl.fB`), nunca por número fijo.
- Anclar charts con `Placement=xlMove` (se mueven con su celda).
- El auditor que tenga referencias por etiqueta, no hardcodeadas (si no, leerá la fila equivocada tras el colapso).

## Render para revisar (charts flotantes)
- `d.Range("A{r1}:AE{r2}").CopyPicture(1,2)` → `ImageGrab.grabclipboard()` SÍ captura charts flotantes que solapan el rango. `export`/`AddChart2` headless dan PNG vacío. Hacer `ScreenUpdating=True`, `Activate()`, `Zoom=80-90`, `time.sleep` entre grabs, reintentar.
- Capturar por secciones (banda a banda) para revisar layout + colores con los ojos.
