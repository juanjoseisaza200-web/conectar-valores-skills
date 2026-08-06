# Matriz de escenarios estándar — 16 sensibilidades

Set por defecto. Cubre los drivers principales de un modelo project finance: indexación (IPC, IPP), tasa de interés (IBR), descuento (Ke), costos (OPEX), y combinados.

## Calibración Fisher Taylor (Colombia)

**No usar Fisher 1:1 ingenuo en shocks grandes.** En realidad Banco de la República tiene reacción **asimétrica**: β=1.0-1.5 al subir IPC, β=0.6-0.8 al bajar (cauto por ZLB).

| IPC shock | IBR Δ acompañante | Factor IPP-IPC Δ |
|---|---|---|
| ±1pp | ±1.0pp / ±0.7pp (-) | 0 |
| ±2.5pp | +3.5pp / -1.5pp | +0.05 / -0.05 |
| ±5pp | +7.5pp / -3.0pp | +0.15 / -0.15 |

Aplicar estas magnitudes en escenarios "Fisher" coordinados. Para shock IPC puro (sin IBR) o IBR puro (sin IPC) NO aplica el coupling.

## Las 16 estándar

| # | Escenario | Shock primario | Shocks acompañantes (coherencia) | Tipo |
|---|---|---|---|---|
| 1 | **IPC +1% (Fisher)** | IPC +1pp | IBR +1pp (Fisher) | Macro |
| 2 | **IPC -1% (Fisher)** | IPC -1pp | IBR -0.7pp (Taylor cauto) | Macro |
| 3 | **Spread IPP +1%** | Factor IPC→IPP +0.01 | — | Sectorial |
| 4 | **Spread IPP -1%** | Factor IPC→IPP -0.01 | — | Sectorial |
| 5 | **IBR +1%** (puro) | IBR +1pp | sin tocar IPC (prima riesgo) | Financiero |
| 6 | **IBR -1%** (puro) | IBR -1pp | — | Financiero |
| 7 | **IBR +2.5%** | IBR +2.5pp | — | Stress crediticio |
| 8 | **IBR -2.5%** | IBR -2.5pp | — | Upside crediticio |
| 9 | **Ke +1% (DDM)** | Rf +1pp → Ke +1pp | — | Valoración |
| 10 | **Ke -1% (DDM)** | Rf -1pp → Ke -1pp | — | Valoración |
| 11 | **Ke +2.5% (DDM)** | Rf +2.5pp | — | Stress descuento |
| 12 | **Ke -2.5% (DDM)** | Rf -2.5pp | — | Upside descuento |
| 13 | **OPEX +10%** | OPEX × 1.10 | — | Operativo |
| 14 | **OPEX -10%** | OPEX × 0.90 | — | Eficiencia |
| 15 | **Downside combinado** | IPC+1%, IBR+1%, Ke+1%, OPEX+10% | Stress integral Fisher | Stress |
| 16 | **Upside combinado** | IPC-0.5%, IBR-0.5%, Ke-1%, OPEX-10% | Best case Fisher | Best case |

## Lógica económica

- **Fisher**: cuando IPC sube e IBR sube juntos, los flujos reales y nominales se reindexan en paralelo. Resultado en equity ≈ 0.
- **IPP independiente del IPC** (en modelos colombianos): IPP indexa tarifa de transmisión. Si IPP sube relativo al IPC, ingresos crecen más que costos → equity sube.
- **IBR puro** (sin IPC): es prima de riesgo crediticio. Sube → más servicio deuda → equity baja.
- **Ke**: tasa de descuento DDM. Sube → menor VP de dividendos → equity baja. Convexidad positiva (bajadas valen más que subidas iguales).
- **OPEX**: directo. Sube → menos margen → equity baja.

## Auditoría de coherencia (CRITICAL)

Cada par +/- debe cumplir:

| Check | Esperado |
|---|---|
| **Simetría** | `\|Δ%(+)\| ≈ \|Δ%(-)\|` con leve convexidad para Ke |
| **Direcciones** | IPP↑→Eq↑, IBR↑→Eq↓, Ke↑→Eq↓, OPEX↑→Eq↓ |
| **Fisher** | shocks IPC+IBR juntos → impacto cercano a 0 |
| **Ranking** | Ke > IBR > OPEX > IPP > IPC Fisher (típico DDM) |

**Si algún par muestra ambos signos negativos cuando debería haber simetría → workflow está mal.**

## Variantes posibles (custom)

Si el usuario quiere algo distinto, posibles shocks comunes:

### Shocks de tarifa
- **Tarifa fija Ecopetrol -10%**: Inputs_C!J99 *= 0.90
- **Tarifa Otrosi #5 +5%**: Inputs_C!J104 *= 1.05

### Shocks de plazo
- **Sin extensión Ecopetrol**: J45 = 0 (apaga flag opciones)
- **Con extensión Ecopetrol**: J45 = 1
- **Plazo mayor +2 años**: J42 += 2 años (DATE math)

### Shocks de capacidad
- **Sin Chivor II 88 MW**: J37 = 0
- **Sin incremento 30 MW**: J32 = 0

### Shocks WACC (no DDM)
- **WACC +50bps**: Beta o ERP en Valoración row 23-27

### Shocks FX
- **+5% USD/COP**: depende del modelo, buscar en Inputs_C tasa de cambio (típicamente row 775-790)

## Matriz REDUCIDA (8 escenarios — para corrida rápida)

Si el usuario quiere correr menos:

1. IPC +1% Fisher
2. IPC -1% Fisher
3. Ke +1%
4. Ke -1%
5. IBR +1%
6. IBR -1%
7. Downside combinado
8. Upside combinado

## Matriz EXTENDIDA (24+ escenarios — para análisis profundo)

Las 16 estándar + custom específicos del activo (tarifa, plazo, capacidad). Pedir al usuario si quiere esto.

---

**Nota:** Los pares siempre se corren en orden +/- consecutivos para facilitar la auditoría visual.
