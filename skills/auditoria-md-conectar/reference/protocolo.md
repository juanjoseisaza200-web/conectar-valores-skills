# Protocolo completo de auditoría MD (metodología de referencia)

Fuente: prompt de auditoría del equipo Conectar (jun-2026) + práctica de mandatos reales. Este documento es la METODOLOGÍA DETALLADA; el flujo operativo con gates está en SKILL.md.

## Severidades
- **Crítico**: cambia la decisión (inversión/deuda/DSCR/covenant/valoración).
- **Alto**: afecta outputs relevantes o credibilidad.
- **Medio**: corregir, no cambia la decisión por sí solo.
- **Bajo**: formato/trazabilidad.
- **Observación**: requiere soporte o discusión (no necesariamente error).

## Fase 1 — Estructura
Separación inputs/cálculo/outputs; colores consistentes; supuestos centralizados; unidades y periodicidad claras; flujo inputs→operación→EEFF→deuda→outputs; checks; auditable por un tercero. RED FLAGS: inputs regados, hojas ocultas sin explicar, hardcodes en fórmulas, valores pegados en proyección, links externos no documentados, fórmulas kilométricas, sin checks, signos inconsistentes.

## Fase 2 — Mecánica
Errores (#REF/#VALUE/#DIV/0/#N/A/#NAME); fórmulas inconsistentes por fila/columna (auditar por R1C1); hardcodes dentro de rangos; fórmulas que se rompen/saltan periodos/no llegan al final; circulares; named ranges; macros; números incrustados (`*1.19`, `/12`, `+500`) sin celda input; signos; diferencias hist vs proyección en la misma fila. Probar reactividad: cambiar un input clave → ¿outputs se mueven con lógica?

## Fase 3 — Inputs
Por supuesto relevante: valor, fuente, fecha, contractual/histórico/benchmark/mgmt/calculado/no soportado, real vs nominal, moneda, antes/después de impuestos, IVA, periodicidad, conservador-razonable-agresivo. Lista mínima: precio/tarifa, volumen/capacidad, inflación, FX, tasas y spreads, capex, opex, márgenes, días CxC/CxP/inventario, impuestos, dividendos, amortización, refi, vida útil, COD, vencimientos contractuales, renovaciones, valor terminal, WACC, múltiplos, degradación/disponibilidad.

## Fase 4 — Histórico vs proyección (CONTRA FUENTE PRIMARIA)
Comparar contra contabilidad real cuenta por cuenta: crecimientos, márgenes (bruto/EBITDA/neto), opex %, capex, WK (días CxC/CxP), endeudamiento, conversión EBITDA→caja, impuestos pagados, dividendos. Calcular: CAGR hist vs proy, a/a, promedios, EMPALME último real → primer proyectado, expansión de márgenes, saltos por cuenta. UMBRALES guía (con juicio): empalme >±10% sin explicación; margen >±200-300 pb; CAGR proy > hist sin driver; WK mejorando >10-15 días; opex < inflación; capex mantenimiento < histórico. RED FLAGS: mejora proyectada que nunca ocurrió, primer año proyectado con salto material, impuestos proyectados ≪ históricos, dividendos agresivos con deuda alta.

## Fase 5 — Backtesting
¿Los drivers proyectados explican el histórico? ingresos=P×Q contra libro (precio implícito = ingreso libro / cantidad real); costos variables se mueven con volumen; fijos son fijos; WK reproduce comportamiento; impuestos consistentes; deuda/intereses reconciliables con saldos×tasas; capex↔PP&E. Identificar sesgos sistemáticos (sobreestima ingresos / subestima opex-capex-WK) y one-offs históricos a normalizar. Concluir: ¿calibrado?, ¿proyecta un desempeño que nunca ocurrió?, ¿mejora soportada o artificial?

## Fase 6 — Conciliaciones mínimas
Balance cuadra · caja flujo = caja balance · utilidad→patrimonio · PP&E: inicial+capex−dep±bajas=final · Deuda: inicial+desembolsos−amort+capitalizados±FX=final · Patrimonio: inicial+aportes+utilidad−dividendos±otros=final · WK: saldos por rotaciones y ∆ con signo correcto en flujo · causado vs pagado (impuestos e intereses) · S&U de la transacción.

## Fases 7-9 — Ingresos / Costos / Capex
INGRESOS: fuentes, contratos y vigencias, moneda, P×Q recalculado, capacidad física/contractual como techo, ingresos post-vencimiento, renovaciones soportadas, indexación (sin doble inflación), one-offs como recurrentes, partes relacionadas, concentración. COSTOS: fijo vs variable, crecimiento vs inflación, eficiencias con soporte, mantenimiento mayor, fees estructura (fiducia/agente/seguros/auditoría), one-offs eliminados pero recurrentes. CAPEX: desglose, curva, IVA, contingencia, IDC, mantenimiento/reposición suficiente vs crecimiento, reconciliación con PP&E, benchmark, sensibilidad sobrecostos 5-20%.

## Fases 10-11 — Deuda / CFADS / DSCR / Covenants
Deuda: S&U, gearing, fechas, moneda, tasa base+spread y periodización, day count, saldo usado para intereses (inicial/promedio/final), gracia, sweep, mini-perm/refi, DSRA (¿sobre el servicio TOTAL?), fees completos, lock-up, tail vs contratos. RECALCULAR: intereses periodo a periodo; roll-forward; **DSCR independiente** = CFADS / servicio TOTAL (intereses + amortización obligatoria + fees recurrentes + sweep si el covenant lo incluye) y comparar contra el reportado — toda exclusión (un tramo, fees) es hallazgo. CFADS = EBITDA − impuestos caja ± ∆WK − capex mantenimiento (± permitidos); NO incluye: caja inicial, deuda/equity nuevo, no recurrentes, efectos sin caja, refi, valor terminal, liberaciones de reservas (si el covenant no lo permite). Revisar DSCR mínimo/promedio/por periodo, periodo más débil y POR QUÉ, headroom, lock-up y default levels, LLCR/PLCR (con guard de cola), DN/EBITDA (¿EBITDA de covenant = EBITDA del modelo?), caja mínima, DSRA requerido vs disponible. RED FLAGS: DSCR con EBITDA en vez de CFADS, caja inicial en numerador, servicio incompleto, covenant anual cuando el contrato es semestral, cumplimiento con metodología incorrecta.

## Fase 12 — Fiscal
Base gravable recalculada; contable vs fiscal; depreciación fiscal; pérdidas y compensación; subcapitalización/límites de intereses; retenciones (¿se recuperan?); IVA recuperable vs costo; ICA/GMF; WHT dividendos e intereses; diferido con contrapartida; tasa efectiva hist vs proy (desviación = hallazgo); causado vs pagado SIEMPRE separados.

## Fase 13 — Working capital
Días hist vs proy CONTRA LIBRO; mejoras de recaudo soportadas; partes relacionadas separadas de comerciales; cartera vencida tratada como cobrable; movimiento = ∆ de saldos del balance; signos; % fijo sin lógica = flag; cuidado con "mejoras" de DSO por mezcla contable (castigos) — distinguir recaudo real de limpieza de bruto.

## Fase 14 — EEFF
P&L/Balance/CF consistentes; checks obligatorios: balance, caja CF=balance, utilidad→patrimonio, deuda↔módulo, PP&E↔capex/dep, WK↔balance, dividendos reducen caja y patrimonio, impuestos e intereses conectan los tres estados. Caja como plug = red flag.

## Fase 15 — Valuation / Returns
Project IRR sin deuda; equity IRR con aportes y distribuciones reales; no mezclar levered/unlevered ni monedas ni real/nominal; valor terminal soportado y % del valor; deuda neta correcta (caja restringida fuera); TIR del lender vs cotización; dividendos solo si cumplen covenants. Si el modelo es de capacidad de deuda (sin valuation): declararlo, no inventar.

## Fase 16 — Sensibilidades
¿Conectadas de verdad? (cambiar input → output); casos no hardcodeados; downside que toque TODOS los módulos; severe downside realmente severo; breakevens (DSCR=covenant, IRR=0, EBITDA, tasa, capex, ingresos); mínimo: base, downside, severe, lender, equity, breakeven. **Pregunta obligatoria: ¿existe sensibilidad del riesgo #1 del negocio?** (en aseo: recaudo/morosidad; en energía: precio/recurso; en concesiones: tráfico/demanda). Si no existe → Alto.

## Fase 17 — Lógica comercial y crediticia
Qué genera el ingreso, quién paga, qué tan contractual, contraparte, concentración, vencimientos vs deuda, tail, ¿se paga con caja operativa real?, ¿depende de refi?, capacidad del sponsor, flexibilidad de costos, distribuibilidad de la caja, holgura de covenants, retorno vs riesgo, qué rompe la estructura, principal riesgo para lender y para equity, supuesto menos soportado.

## §4 — Las 25 preguntas MD (responder TODAS al final)
1 drivers top-5 · 2 periodo de mayor estrés · 3 DSCR mínimo y por qué · 4 supuesto más agresivo · 5 menos soportado · 6 mayor impacto en IRR · 7 en DSCR · 8 en deuda máxima · 9 ¿se paga con caja real? · 10 ¿depende de refi? · 11 ¿deuda vence antes que contratos? · 12 ¿tail suficiente? · 13 ¿capex completo? · 14 ¿opex realista? · 15 ¿ingresos soportados? · 16 ¿proyección consistente con histórico? · 17 ¿márgenes defendibles? · 18 ¿WK razonable? · 19 ¿impuestos bien modelados? · 20 ¿valoración bien calculada? · 21 ¿sensibilidades útiles? · 22 ¿red flags que paran comité? · 23 ¿qué corregir antes de circular? · 24 ¿qué información falta? · 25 conclusión MD.

## §5 — Formato del entregable (20 secciones)
1 Executive Summary · 2 Confiabilidad (Confiable/Con ajustes/Parcial/No confiable/No apto) · 3 Top 10 hallazgos (tabla) · 4 Model Map · 5 Mecánica · 6 Supuestos · 7 Hist vs proy (+anexo xlsx) · 8 Backtesting · 9 Conciliaciones · 10 Ingresos · 11 Costos/OPEX · 12 Capex · 13 Deuda y covenants · 14 Fiscal · 15 Valuation/returns · 16 Sensibilidades y downside · 17 Red flags de comité · 18 Recomendaciones priorizadas (celda+esfuerzo) · 19 Información faltante · 20 Conclusión final MD ("¿lo usaría hoy? ¿qué corregir antes?") + Anexo 25 preguntas.

## Profundidad exigida
Nunca "parece razonable": qué revisaste, qué encontraste, hoja/celda, por qué importa, impacto cuantificado, cómo se corrige, qué soporte falta. Nunca "crece mucho": cuánto, contra qué histórico, qué cuenta, qué margen, con o sin justificación, qué sensibilidad correr. Nunca "el DSCR baja": en qué periodo, por qué (ingresos/opex/capex/tasa/amortización/impuestos/WK), si incumple covenant, si hay caja, si es estructural o temporal.
