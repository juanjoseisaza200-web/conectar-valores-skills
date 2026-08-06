# Razonar QUÉ presentar (FASE 1)

Un dashboard malo vuelca números; uno bueno **responde las preguntas de quien decide**. Antes de construir, razona a profundidad. Esto NO es un layout fijo a copiar: es un marco para pensar.

## 0. Encuadre (responder primero)
- **¿Qué transacción es?** (deuda project-finance, adquisición, refi, equity…). Define qué importa.
- **¿Quién lee?** Lender/banco → cobertura, covenants, liquidez, downside. Comité de crédito → lo mismo + estructura. Inversionista equity → retornos, crecimiento, márgenes. Junta → ejecutivo + operativo.
- **¿Cuáles son los 3 riesgos clave?** (del modelo/auditoría: ej. recaudo/morosidad, liquidez ajustada, año débil de DSCR). El dashboard debe hacerlos VISIBLES, no esconderlos.
- **Unidades y horizonte**: COP mil MM (=modelo/1e6), años 2026-2039 (cols Q:AD del modelo).

## 1. Las 10 dimensiones (recorrerlas TODAS; elegir lo relevante de cada una)
Para cada métrica anota su **`Hoja!celda`** (de FASE 0). Si el modelo no la calcula: o se construye con fórmula trazable, o se marca NO disponible — nunca se inventa.

1. **Ejecutivo (banda KEY INFORMATION, 8 tarjetas)** — lo que un MD mira en 5 segundos: EBITDA y margen, utilidad neta, TIR all-in, DSCR mínimo, Deuda Neta/EBITDA pico, deuda bruta, liquidez (meses OPEX), y una más del riesgo clave. Número grande + sub (▲▼ vs covenant/año).
2. **P&L / resultados** — ingresos, EBITDA, margen EBITDA, EBIT, utilidad neta, margen neto. Chart: ingresos+EBITDA barras + margen línea (eje 2º). Puente D&A si aplica.
3. **Retornos** — ROE, ROA, payout, patrimonio. Chart: utilidad neta barras + ROE/ROA líneas.
4. **Deuda / cobertura / covenants** (núcleo de un deal de deuda):
   - Saldos por tramo (paydown, área apilada), amortización + intereses, servicio de deuda.
   - **DSCR** (¡ojo a la definición! a veces el covenant excluye un tramo → reportar el DSCR con servicio TOTAL, y aparte el del covenant), ICR, **DSCR mín real vs covenant**.
   - WAL por tramo (verificar el blended: debe estar entre el de cada tramo, ponderado por saldo; si el modelo da uno fuera de rango es BUG → recalcular en el dashboard).
   - DSRA (target meses vs fondeado), comisión, TIR all-in por tramo vs cotización de mercado (NETGO).
   - Flag de cumplimiento de covenants.
5. **Apalancamiento** — Deuda Neta/EBITDA vs covenant (≤X), headroom, gearing, Deuda Bruta/EBITDA, estructura de capital (deuda vs patrimonio), % amortizado acumulado.
6. **Liquidez (suele ser riesgo clave)** — caja, IRA, DSRA, liquidez total, **meses de OPEX**, DSO. Si la liquidez es ajustada (<1 mes) hazla protagonista, no la diluyas.
7. **Working capital** — Δ WC, DSO/DPO/DIO, ciclo de caja, cartera comercial vs intercompany.
8. **Operativo (drivers físicos)** — volúmenes (toneladas, km, suscriptores), productividad (ingreso/unidad, costo/unidad), ingreso por línea de servicio (mix), por región.
9. **Calidad de cartera / riesgo** — morosidad (nivel y tasa), provisión acumulada, cobertura provisión/cartera, castigo. Tendencia (¿mejora o empeora?).
10. **Flujo de caja (cascada)** — EBITDA → (−impuestos −capex ±WC) → CFADS → (−servicio) → superávit a equity → dividendos. Tabla año a año + bridge acumulado. **Debe FOOTAR**: si la suma de la tabla no cuadra con CFADS del modelo, agregar fila conciliadora "(+/−) Otros" (= CFADS − componentes), no dejar el descuadre.

## 2. Reglas de selección (criterio MD)
- **Trazabilidad total**: cada cifra ancla a una celda. Cero alucinación.
- **Sin redundancia**: un dato se grafica UNA vez en su lugar más lógico (hero arriba). Si un slot quedaría repetido, reemplázalo por OTRO dato relevante no graficado (no lo borres dejando hueco, ni lo dupliques).
- **Charts que cuentan historia > comparaciones planas**: una serie temporal que baja/sube (desapalancamiento, cobertura que mejora, paydown) comunica; 3 barras casi iguales (ej. TIR vs cotización por tramo) no dicen nada → prefiere tendencia, composición o evolución.
- **Hacer visible el riesgo**: si el deal tiene un punto débil (año de DSCR<1.0x pre-IRA, liquidez <0.5 mes, morosidad estresada no modelada), el dashboard lo muestra con su semáforo/anotación. Asesorar ≠ maquillar.
- **Honestidad sobre límites del modelo**: si LLCR≡PLCR (el modelo no separa vida-préstamo de vida-proyecto), si un covenant está mal definido, si una serie es plana por ser saldo de balance — anótalo/relabel en vez de fingir.
- **Conservar todo, hacerlo gráfico**: no eliminar métricas valiosas; convertir tablas densas en charts y agregar las que falten.
- **Mostrar todos los años o un set claro**: la tabla compacta C:P lleva los 14 años; las tarjetas KPI muestran el valor del año relevante (2026 o el pico/mínimo del horizonte según la métrica).

## 3. Salida de la fase
El **plan de contenido** (`templates/plan-contenido.md`): por sección, sus filas/KPIs y charts, cada uno con `Hoja!celda`. Es lo que se aprueba en GATE 1 y lo que guía el script de FASE 3. Si una fuente no se encontró en FASE 0, vuelve a escanear — no construyas con un hueco.
