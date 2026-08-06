# Plan de contenido del dashboard — <PROYECTO>

> Salida de FASE 1. Cada cifra ancla a `Hoja!celda` (verificada en FASE 0). Se aprueba en GATE 1.
> Encuadre: deal=<tipo> · audiencia=<lender/comité/inversionista> · unidades=COP mil MM (=modelo/1e6) · horizonte=<2026-2039> · riesgos clave=<r1, r2, r3>.

## Banda KEY INFORMATION (8 tarjetas)
| # | Título | Fórmula (fuente) | Formato | Sub |
|---|---|---|---|---|
| 1 | EBITDA <año> | `=Waterfall!Q60/1e6` | `#,##0.0` "mil MM" | ▲ vs año ant. |
| 2 | Margen EBITDA | `=Waterfall!Q60/EEFF!Q289` | `0.0%` | |
| 3 | Utilidad neta | `=EEFF!Q452/1e6` | mil MM | |
| 4 | TIR all-in | `=Deuda!F513` | `0.0%` | vs NETGO +Xpb |
| 5 | DSCR mín | `=Deuda!F557` | `0.00"x"` | ≥ covenant |
| 6 | DN/EBITDA pico | `=Deuda!F547` | `0.00"x"` | ≤ covenant |
| 7 | Deuda bruta | `=Deuda!Q540/1e6` | mil MM | |
| 8 | Liquidez | `=...` meses OPEX | `0.0` "meses" | riesgo clave |

## Secciones (por cada una: filas compactas + charts, con fuente)
### <SECCIÓN> (banda navy)
- **Tabla compacta** (col B etiqueta, C:P años):
  - `<métrica>` ← `Hoja!{Q..AD}fila` (sign, scale)
  - …
- **Chart(s)** (R:AD): tipo, título, series=[(nombre, fila compacta C:P, color), …]
  - Ej.: ColCluster "Ingresos · EBITDA · Margen" = [(Ingresos, r69, TEAL),(EBITDA, r71, NAVY),(Margen, r72, GOLD, line2)]

### Repetir por: Operación · Rentabilidad/Costos · Cobertura del servicio de deuda · Apalancamiento · Liquidez · Concentración · Macro · Operativos · Resumen Financiero · Resumen de Deuda · Panel Operativo · Calidad de Cartera

## Cascada de caja (final)
Pasos (cada uno con su fila/fuente): EBITDA → (−)Impuestos → (−)CAPEX → (+/−)WC → [(+/−)Otros conciliador] → **CFADS** → (−)Servicio → **Superávit a equity** → (−)Dividendos.
- Verificar FOOTING: suma de componentes = CFADS del modelo; si no, fila "Otros" = CFADS − componentes.

## Límites del modelo a declarar (no fabricar)
- <ej. LLCR≡PLCR: el modelo no separa vida-préstamo de vida-proyecto → relabel honesto>
- <ej. DSCR del covenant excluye un tramo → mostrar DSCR con servicio total + nota>

## Interactividad
- Filtro de periodo (Desde/Hasta) — sí/no.
- Escenarios: vía `Inputs_C!F6` (manual) — los charts reflejan el escenario activo al recalcular.
