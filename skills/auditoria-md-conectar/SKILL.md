---
name: auditoria-md-conectar
description: Auditoria exhaustiva de modelos financieros con mentalidad de Managing Director (model audit, project finance, debt advisory) - estandar Conectar Valores. Usar SIEMPRE que se pida auditar, revisar o validar un modelo financiero Excel para comite de inversion/credito, lenders o inversionistas. Protocolo de 17 fases con gates y herramientas ejecutables: integridad mecanica, supuestos, historico-vs-proyeccion CONTRA LA FUENTE PRIMARIA (balance de prueba), backtesting PxQ, conciliaciones, ingresos/costos/capex, deuda/CFADS/DSCR/covenants, fiscal, working capital, EEFF, retornos, sensibilidades y logica crediticia. Disenada para ejecutarse sin errores por cualquier modelo (Fable/Opus/Sonnet).
---

# Auditoría MD de modelos financieros — estándar Conectar

**Mentalidad**: MD revisando el trabajo de un analista antes de presentarlo a un banco/fondo/comité. Escéptico, granular, accionable. NO asumas que el modelo está bien; NO confíes en outputs sin recalcular cómo se producen; NO aceptes supuestos sin fuente. **Solo lectura**: jamás modificar el archivo auditado; toda corrección se indica con hoja/celda.

Antes de empezar: leer `reference/protocolo.md` (metodología completa, umbrales, red flags por fase, severidades, 25 preguntas MD) y `LECCIONES.md` de `aprender-de-errores`. Si el modelo es del estándar Conectar, usar también las herramientas de `modelaje-conectar/tools/`.

## Reglas duras (no negociables)

1. **AUDITORÍA = DESDE CERO.** No reutilizar conclusiones de sesiones previas sin re-verificarlas contra el archivo VIGENTE; si se reutiliza una verificación anterior, declararlo explícitamente con fecha y alcance.
2. **La fase histórico-vs-proyección se corre CONTRA LA FUENTE PRIMARIA** (balance de prueba/contabilidad), cuenta por cuenta, con la herramienta — no con agregados de memoria.
3. Todo hallazgo lleva: severidad (Crítico/Alto/Medio/Bajo/Observación), hoja/celda, por qué importa, impacto cuantificado, recomendación con fórmula/celda, soporte faltante.
4. Outputs de ratios (DSCR, LLCR, leverage) se RECALCULAN independientes de los del modelo y se comparan — la diferencia es hallazgo.
5. Supuesto sin fuente = "requiere soporte" aunque parezca razonable. Output correcto con lógica económica débil = hallazgo. Fórmula que funciona pero no es auditable = hallazgo.
6. No declarar la auditoría completa sin el GATE de cobertura (abajo) en verde.

## FASES (resumen ejecutivo — detalle y red flags por fase en `reference/protocolo.md`)

| # | Fase | Herramienta / técnica | Gate |
|---|---|---|---|
| 0 | Preparación | Copia de solo lectura; mapear hojas/nombres/macros/escenarios | Model Map completo |
| 1 | Estructura | Inspección + convenciones (inputs centralizados, colores, flujo) | — |
| 2 | Mecánica | `modelaje-conectar/tools/auditar_modelo.py` (errores, circulares, FAST, checks) + scan hardcodes embebidos | 0 errores explicados |
| 3 | Inputs/supuestos | Tabla: valor, fuente, fecha, clasificación (contractual/histórico/benchmark/mgmt/no soportado), conservador-razonable-agresivo | Tabla completa de los inputs top-20 |
| 4 | **Hist vs proy** | **`tools/hist_vs_proy.py` contra el balance de prueba** (empalme, CAGRs, flags ±10%/300pb) | Anexo xlsx generado |
| 5 | Backtesting | PxQ: precio implícito libro vs modelo; cantidades vs fuentes; sesgos sistemáticos | Conclusión calibración |
| 6 | Conciliaciones | Balance, caja flujo=balance, deuda/PP&E/WK/patrimonio roll-forwards, causado vs pagado | Tabla 9 conciliaciones |
| 7-9 | Ingresos/Costos/Capex | Recalcular desde drivers; capacidad; vencimientos; fijo-vs-variable; capex↔PP&E; mantenimiento | — |
| 10-11 | Deuda/CFADS/DSCR/covenants | `tools/capturar_outputs.py` + **recalcular DSCR con servicio TOTAL** y comparar con el reportado; CFADS componente a componente; DSRA; tail; refi | DSCR recalculado vs reportado |
| 12 | Fiscal | Base gravable, causado vs pagado, pérdidas, retenciones, diferido, tasa efectiva hist vs proy | — |
| 13 | Working capital | Días hist vs proy CONTRA LIBRO; partes relacionadas separadas; signos en flujo | — |
| 14 | EEFF | Checks obligatorios (lista en protocolo) | — |
| 15 | Valuation/returns | Si aplica: IRR project/equity sin mezclar flujos; TIR lender vs cotización | — |
| 16 | Sensibilidades | ¿Conectadas de verdad? cambiar input → output; downside severo; breakevens; ¿existe el estrés del RIESGO #1 del negocio? | — |
| 17 | Lógica crediticia | Preguntas MD; periodo más débil y por qué; ¿se paga con caja real?; tail; liquidez en meses de OPEX | 25 preguntas respondidas |

## GATE DE COBERTURA (antes de entregar)

Checklist de completitud — todo SÍ o declarado como limitación con razón:
- [ ] Fase 4 corrida contra fuente primaria con anexo generado (no de memoria)
- [ ] DSCR/covenants recalculados independientemente y comparados
- [ ] CFADS descompuesto y validado componente a componente
- [ ] Liquidez expresada en meses de OPEX
- [ ] Política de dividendos vs lock-up/caja mínima revisada
- [ ] Tenor/WAL vs term sheet
- [ ] Sensibilidad del riesgo #1 del negocio existe (si no: hallazgo Alto)
- [ ] Hallazgos previos de auditorías anteriores: re-verificados y marcados [VIGENTE]/[RESUELTO]
- [ ] Cada Crítico/Alto tiene celda + impacto cuantificado + corrección concreta
- [ ] 25 preguntas MD respondidas; veredicto final con condiciones explícitas

## ENTREGABLES (formato en `reference/protocolo.md` §5)

1. Documento de 20 secciones (Executive Summary → Conclusión MD) como .md en la carpeta del proyecto; Word VF3 con `conectar-docx-creator` si el usuario lo pide.
2. Anexo xlsx hist-vs-proy (lo genera la herramienta).
3. Audit log con columnas: Severidad | Categoría | Hoja | Celda | Hallazgo | Por qué importa | Impacto | Recomendación | Prioridad.
4. Veredicto: Confiable / Confiable con ajustes / Parcialmente confiable / No confiable / No apto para comité — con el "¿lo firmaría hoy?" explícito.

## Herramientas

- `tools/hist_vs_proy.py` — motor hist-vs-proyección contra balance de prueba (editar el MAPA PUC↔filas del modelo al inicio del script).
- `tools/capturar_outputs.py` — captura series de outputs (DSCR, covenants, caja, liquidez) de un modelo en solo lectura.
- `modelaje-conectar/tools/auditar_modelo.py` y `cv_model.py` — mecánica, checks, circulares, escenarios.
- Reglas COM (esperar CalculationState, DispatchEx, etc.): `modelaje-conectar/reference/com-y-escenarios.md`.
