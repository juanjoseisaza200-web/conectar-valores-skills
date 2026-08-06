# Outputs configurables — Qué celdas podés trackear

El script soporta **cualquier número de outputs** (cualquier celda de cualquier hoja). Por cada output, el reporte muestra valor + Δ% vs baseline.

## Cómo configurar

### Opción A: vía prompt inicial (recomendado)

En la sección "Outputs a trackear" del prompt inicial, lista los que querés. Claude edita el script automáticamente.

```
Outputs:
- Equity DDM (Valoración!F137)
- DSCR mínimo (Financing!F250)
- Deuda 2030 (Financing!AE65)
- FCF total (Valoración!F54)
- EBITDA 2026 (EEFF IFRS!K50)
```

### Opción B: editar `$OUTPUTS` en `run_sensibilidades.ps1`

```powershell
$OUTPUTS = @(
    @{label="Equity DDM";   sheet="Valoración"; cell="F137"; type="value"},
    @{label="Deuda Neta";   sheet="Valoración"; cell="F123"; type="value"},
    @{label="DSCR min";     sheet="Financing";  cell="F250"; type="value"},
    @{label="FCF total";    sheet="Valoración"; cell="F54";  type="value"}
)
```

Cada entry necesita:
- `label`: nombre de columna en hoja Resultados
- `sheet`: nombre exacto de la hoja
- `cell`: notación A1
- `type`: `value` (default) o `pct` (formato porcentaje)

## Catálogo de outputs típicos Conectar

### Valoración

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| Equity DDM | Valoración | F137 | Métrica principal DDM |
| Equity WACC | Valoración | F124 | Bridge desde EV via WACC |
| Equity Pagos Restringidos | Valoración | F148 | DDM solo PR |
| EV WACC | Valoración | F117 | Sum VP FCFF |
| Deuda Neta (Dic 2025) | Valoración | F123 | F121 - F122 |
| Promedio 3 métodos Equity | Valoración | F159 | (WACC + DDM + PR) / 3 |
| WACC promedio | Valoración | F107 (anual) | varía por periodo |
| Ke anual | Valoración | F86 | varía por periodo |

### Deuda (Financing)

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| Saldo deuda final | Financing | KP65 (o última col) | Debe ser ~0 si amortiza |
| Saldo deuda año X | Financing | columna del año, fila 65 | Buscar año específico |
| Servicio deuda total | Financing | sumar filas 656-699 | Intereses + capital |
| DSCR mínimo | Financing | buscar "DSCR" en col E | Ratio cobertura |

### Flujos (Valoración)

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| FCFF total | Valoración | F54 (sum) | Sum row 54 J:KP |
| FCFE total | Valoración | F131 (sum) | Para DDM |
| VP(FCFF) total | Valoración | F117 | = EV WACC |
| Pagos Restringidos total | Valoración | F148 | Equity PR |

### Operacional (EEFF IFRS / Ingresos / Opex)

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| Ingresos año X | Ingresos | columna del año, sumar | Buscar año |
| EBITDA año X | EEFF IFRS | columna del año | Buscar "EBITDA" en col E |
| OPEX total año X | Opex | columna del año, fila 248 | Total Opex |
| Utilidad neta año X | EEFF IFRS | buscar "Utilidad neta" | |
| FCO año X | EEFF IFRS | buscar "Flujo Caja Operativo" | |

### Equity

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| Dividendos pagados total | Equity | sumar fila relevante | |
| Aportes de capital | Equity | buscar "Aportes" | |
| Patrimonio final | EEFF IFRS | última col, fila Patrimonio | |

### Tax

| Output | Hoja | Celda | Notas |
|---|---|---|---|
| Total impuestos pagados | Tx | sumar fila relevante | |
| t_eff promedio | Tx | promedio fila t_eff | |

## Ejemplo: análisis completo de balance entre métodos

Si querés ver cómo cambian las 3 métricas de equity (WACC vs DDM vs PR) en cada escenario:

```powershell
$OUTPUTS = @(
    @{label="Eq WACC";  sheet="Valoración"; cell="F124"; type="value"},
    @{label="Eq DDM";   sheet="Valoración"; cell="F137"; type="value"},
    @{label="Eq PR";    sheet="Valoración"; cell="F148"; type="value"},
    @{label="Eq Promedio"; sheet="Valoración"; cell="F159"; type="value"},
    @{label="Deuda Neta"; sheet="Valoración"; cell="F123"; type="value"}
)
```

Tabla resultante (5 outputs × 2 cols = 10 cols + 3 base = 13 cols totales).

## Ejemplo: análisis de deuda año a año

```powershell
$OUTPUTS = @(
    @{label="Eq DDM";       sheet="Valoración"; cell="F137"; type="value"},
    @{label="Deuda 2026";   sheet="Financing";  cell="J65";  type="value"},
    @{label="Deuda 2028";   sheet="Financing";  cell="X65";  type="value"},
    @{label="Deuda 2030";   sheet="Financing";  cell="AL65"; type="value"},
    @{label="DSCR min";     sheet="Financing";  cell="F250"; type="value"}
)
```

## Ejemplo: análisis operacional

```powershell
$OUTPUTS = @(
    @{label="Eq DDM";        sheet="Valoración"; cell="F137"; type="value"},
    @{label="EBITDA 2026";   sheet="EEFF IFRS";  cell="J50";  type="value"},
    @{label="EBITDA 2030";   sheet="EEFF IFRS";  cell="AL50"; type="value"},
    @{label="OPEX 2026";     sheet="Opex";       cell="J248"; type="value"},
    @{label="Ingresos 2026"; sheet="Ingresos";   cell="J217"; type="value"}
)
```

## Limitaciones

- **No** outputs que sean SUMs dinámicas — el script lee solo celdas individuales. Si necesitás un total (ej. "FCF total acumulado"), creá una celda en el modelo con el SUM y trackea esa celda.
- **Performance**: cada output adicional añade ~50ms por escenario (lectura COM). Con 10 outputs y 16 escenarios → +8 segundos total. Trivial.
- **Pares simétricos**: la auditoría usa el primer output en `$OUTPUTS` como métrica principal. Pone primero el output más relevante (ej. Equity DDM).

## Cómo encontrar la celda de un output que no conocés

1. Abrir modelo en Excel
2. Buscar (Ctrl+F) el label en columna E de la hoja relevante
3. Anotar `Hoja!Celda`
4. Agregar al `$OUTPUTS` o decírselo al Claude en el prompt
