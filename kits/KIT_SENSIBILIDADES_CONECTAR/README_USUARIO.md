# 🎯 KIT SENSIBILIDADES — Modelos Project Finance Conectar Valores

**Para el usuario.** Lee esto primero. Tiempo total típico: 5 min de setup + ~20-40 min de corrida automática.

---

## ¿Qué hace este kit?

Corre escenarios de sensibilidad sobre cualquier modelo financiero con la estructura típica de Conectar Valores (Inputs_C con MATCH/INDEX, macro `sensibilidad`, valoración DDM). Genera una tabla con 16 escenarios estándar (o los que quieras) y deja todo grabado en una hoja del modelo.

## Pasos para usarlo

### 1. Copia tu modelo a esta carpeta

Pon una copia (no el original) de tu modelo .xlsm aquí en `KIT_SENSIBILIDADES_CONECTAR/`.

### 2. Abre Claude Code (en cualquier computador)

Asegúrate de tener:
- Windows con Excel instalado
- Permisos para correr macros (Trust Center → enable macros)
- PowerShell

### 3. Pega el prompt inicial

Abre `PROMPT_INICIAL.md`, copia el bloque de prompt, **reemplaza el nombre de tu archivo** y pégalo en Claude. Eso es todo. Claude se encarga del resto.

Si el modelo tiene el layout estándar de Conectar (igual mapping de filas), Claude correrá sin preguntarte nada. Si es un modelo distinto, te preguntará SOLO lo que no pueda inferir solo (ej: si la macro se llama distinto, o si el Equity DDM está en otra celda).

### 4. Espera y revisa

Claude te dará updates en tiempo real. Cada escenario tarda 30-120 segundos. Al final, abrirás tu modelo y verás la hoja `Reusltado Sensibilidades` (o el nombre que tenga) con los 16 escenarios.

---

## ¿Qué te puede preguntar Claude?

Solo si NO puede mapear automáticamente:
1. **Nombre exacto del archivo** si pusiste varios
2. **Nombre de la macro** si no se llama `sensibilidad`
3. **Si querés sensibilidades adicionales o distintas** a las 16 estándar
4. **Si hay shocks específicos del activo** (ej: capacidad MW, plazos contrato)

**No te preguntará** (lo infiere o ya lo sabe):
- Estructura de Inputs_C, MATCH/INDEX
- Cómo correr la macro
- Workflow J-directo / K-backup
- Qué celdas leer para Equity DDM
- Cómo manejar errores COM/RPC
- Si hacer auditoría de pares simétricos (siempre la hace)

---

## Las 16 sensibilidades estándar

| # | Shock | Magnitud |
|---|---|---|
| 1-2 | IPC Fisher (con IBR coherente) | ±1% |
| 3-4 | Spread IPP-IPC | ±1% |
| 5-8 | IBR puro | ±1%, ±2.5% |
| 9-12 | Ke (descuento DDM) | ±1%, ±2.5% |
| 13-14 | OPEX | ±10% |
| 15 | Downside combinado | Stress integral Fisher |
| 16 | Upside combinado | Best case Fisher |

Si querés otras (ej: shock de tarifa, plazo contrato, capacidad), simplemente pídeselo a Claude después.

---

## Output esperado

1. **Hoja Resultados** del modelo con tabla de 17 filas (BASELINE + 16) y columnas dinámicas según outputs que pediste:
   - Por defecto: # / Escenario / Notas / Equity DDM / Δ% / Deuda Neta / Δ% / Audit
   - Si pediste más outputs (ej. DSCR, FCF, deuda año X), aparecen cada uno con su columna y Δ%
2. **Resumen en chat** con análisis de pares simétricos y ranking de sensibilidades
3. **Archivo guardado** (con nombre `MODELO_<empresa>_SENSIBILIDADES.xlsm` en la misma carpeta)

## Outputs flexibles (qué podés trackear)

No estás limitado a Equity DDM. Podés pedir cualquier celda de cualquier hoja del modelo:

- **Valoración**: Equity DDM, Equity WACC, Equity PR, EV WACC, etc.
- **Deuda (Financing)**: DSCR mínimo, saldo deuda año X, servicio deuda total
- **Flujo (Valoración o Financing)**: FCF total, FCFE total, pagos restringidos
- **Operacional (EEFF IFRS, Ingresos, Opex)**: EBITDA año X, ingresos año X, OPEX total
- **Equity**: total dividendos, retained earnings
- **Tax (Tx)**: total impuestos pagados, t_eff promedio

Especificás esto en el prompt inicial. Por ejemplo:

> "Trackea: Equity DDM, DSCR mínimo en Financing!F250, deuda 2030 en Financing!AE65, y FCF total en Valoración!F54"

---

## Si algo sale mal

- **"F137 = #N/A" o equity en error**: el modelo está corrupto. Cierra todo Excel, copia tu modelo original de nuevo, y arranca de cero.
- **Claude se queda muchos minutos sin respuesta**: la macro tarda 30-120s por escenario, total 20-40 min para 16 escenarios. Esperar.
- **Excel no abre o falla COM**: cerrar TODOS los procesos Excel (`taskkill /F /IM EXCEL.EXE`) y reintentar.
- **OneDrive sync conflict**: trabajar en la carpeta del Desktop ayuda, pero si persiste, copiar a `C:\Temp\` (fuera de OneDrive).

---

## Archivos en este kit

| Archivo | Para quién | Uso |
|---|---|---|
| `README_USUARIO.md` | TÚ | Estás leyendo esto |
| `PROMPT_INICIAL.md` | TÚ | Copia el prompt y pásalo a Claude |
| `INSTRUCCIONES_CLAUDE.md` | Claude | Auto-instrucciones que Claude lee primero |
| `WORKFLOW_TECNICO.md` | Claude | Playbook técnico detallado |
| `matriz_escenarios.md` | Ambos | Las 16 sensibilidades estándar y variantes |
| `outputs_configurables.md` | Ambos | Cómo trackear cualquier celda (deuda, FCF, DSCR, etc.) |
| `run_sensibilidades.ps1` | Claude | Script PowerShell adaptable que ejecuta |
| `build_organized.py` | Claude | Post-procesador: genera xlsx multi-hoja organizado (Resumen + Tablas + Series mensuales) |

---

## ¿Cuándo NO usar este kit?

- Si el modelo NO tiene macro de sensibilidad (Claude la armará desde cero — distinto workflow)
- Si querés cambiar la estructura del modelo (no es para eso)
- Si los inputs están en hojas no estándar y no te interesa mapear (Claude puede pero requiere tiempo)

Para cualquier otro caso, este kit funciona.
