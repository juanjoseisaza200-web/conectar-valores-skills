# Paleta, tipografía y layout (one-pager)

## Paleta corporativa (valores `.RGB` de COM = empaquetado BGR)
COM devuelve/recibe el color como `B*65536 + G*256 + R`. Usa estas constantes (en `dash_lib`):

| Nombre | Valor COM | RGB | Uso |
|---|---|---|---|
| NAVY  | 6174487  | (23,55,94)    | barras/serie 1, bandas de sección, títulos |
| TEAL  | 8882990  | (46,139,135)  | serie 2 (acento) |
| GOLD  | 4826313  | (201,164,73)  | serie 3, header de tablas numéricas, líneas clave |
| MINT  | 13434828 | (204,229,204) | fila/celda destacada (CFADS, superávit, "CUMPLE") |
| RED   | 2960895  | (63,45,45→rojo)| covenant/piso/alerta (línea de referencia) |
| SLATE | 11507320 | (120,150,175) | serie 4 (apilados de 5) |
| LGRAY | 13157310 | (190,195,200) | serie 5 / filas alternas suaves |
| DGRAY | 6710886  | (102,102,102) | texto secundario, fuente de charts |
| GRID  | 15066597 | gris claro    | gridlines |
| WHITE | 16777215 | blanco        | fondo de chart, relleno de celda |

`dl.rgb(r,g,b)` construye un color nuevo correctamente empaquetado.
**Regla**: serie 1=NAVY, 2=GOLD, 3=TEAL; líneas de covenant/piso=RED; apilados de 4-5 añaden SLATE y LGRAY. Nunca dejar el color default del tema (azul 8544277 / naranja 3305961 / verde 2386713 son señal de serie SIN colorear → corregir).

## Tipografía y números
- Fuente **Arial** en todo. Tamaños: banda de sección 12 bold; sub-header 10 bold; títulos de chart 11 bold NAVY; etiquetas de tabla 8; datos 8; tarjeta KPI número ~20-24 bold NAVY, título 8 bold blanco sobre NAVY, sub 7.
- Formatos: mil MM `#,##0.0`; ratios `0.00"x"`; porcentajes `0.0%`; años `0"a"`; meses `0"m"`. Eje de chart numérico `#,##0` (¡NUNCA heredar `0.0%` de un chart anterior — da "45000%"!).
- Alineación: datos centrados; etiquetas a la izquierda.

## Tabla compacta (convención CLAVE de los charts)
- Los **años 2026-2039** viven en `Dashboard!Q7:AD7` (cols 17-30) — fila de años global.
- Cada fila de métrica espeja el modelo en **C:P** (cols 3-16): `C=2026 … P=2039`, mapeando `C↔Q(modelo), D↔R … P↔AD`. La etiqueta va en **col B**.
- `dl.compact_mirror(sh, row, label, sheet, mrow, sign, scale)` escribe `C{row}=±{sheet}!Q{mrow}/scale … P{row}=±{sheet}!AD{mrow}/scale`.
- Los charts leen **`Values=$C$r:$P$r`, `XValues=$Q$7:$AD$7`**. Los datos originales Q:AD del modelo (si se importan) se ocultan con fuente blanca + col Q angosta.
- Helpers de chart (series no-FAST: por categoría, cascada, mini-tablas) van en la hoja oculta **`_DashViz`**.

## Tarjetas KPI (banda KEY INFORMATION)
- 8 tarjetas a todo el ancho (2 filas × 4). Celdas FUSIONADAS por tarjeta (la grilla angosta corta los números → "##").
- Estructura por tarjeta: título (bold blanco sobre NAVY) / número grande (bold NAVY) / sub (▲▼ + referencia, gris).
- **Las etiquetas NO empiezan con "="** (Excel las toma como fórmula → `#NAME?`). El número sí es fórmula (`=EEFF!Q452/1e6`).
- `dl.kpi_band(sh, start_row, cards)` donde `cards` = lista de `(titulo, formula, fmt, sub)`.

## Bandas y paneles
- **Banda de sección**: `dl.band(sh, row, "TEXTO")` — merge B:AD, fondo NAVY, texto blanco bold 12, alto 20.
- **Sub-header**: `dl.band_sub(sh, row, "texto")` — texto NAVY bold 10 (sin fondo).
- **Panel temático** = banda navy + **tabla compacta C:P a la izquierda** (etiqueta col B, años en cabecera) + **gráfica en R:AD a la derecha** (cols 18-30) que lee la tabla compacta. Filas alternas LGRAY/WHITE.

## Layout one-pager (scroll vertical)
1. **FILTRO DE PERIODO** (fila ~8-9): "Desde [año] / Hasta [año]" (validación en celda).
2. **KEY INFORMATION** (banda + 8 tarjetas).
3. **Gráficas hero** (3 charts clave del deal: cobertura/CFADS·DSCR, paydown por tramo, DSCR vs covenant).
4. **Paneles temáticos** (Operación, Rentabilidad/Costos, Cobertura, Apalancamiento, Liquidez, Concentración, Macro, Operativos) — cada uno tabla compacta + chart.
5. **Secciones de detalle** con charts hero (Resumen Financiero, Resumen de Deuda, Panel Operativo, Calidad de Cartera).
6. **Resumen de Flujo de Caja** (tabla año a año + cascada bridge) al final.

## Geometría de charts
- Posicionar por celda: `co.Left=Columns(c1).Left`, `co.Top=Rows(r1).Top`, ancho/alto por diferencia de columnas/filas (`dl.place_chart`). Charts de panel: tabla en C:P, chart desde col R (18) a AD (30).
- **Anclar TODOS con `Placement=xlMove`** (`dl.anchor_all_charts`) → se mueven con su celda si la hoja se compacta.
- Evitar solapes: cada chart cabe en su banda [banda+1 .. siguiente_banda−1]. El auditor reporta solapes.
- **Sin sparklines.** **Sin data bars** salvo que el usuario los pida.
