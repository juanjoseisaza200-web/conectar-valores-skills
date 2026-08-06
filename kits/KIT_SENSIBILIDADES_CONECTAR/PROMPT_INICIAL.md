# PROMPT INICIAL — Copiar y pegar a Claude

Después de copiar tu modelo (`.xlsm`) a esta carpeta, abre Claude Code en el directorio `KIT_SENSIBILIDADES_CONECTAR/` y pega este prompt. **Reemplaza** los `[CAMPOS]` con tu info real.

---

## Prompt (copiar todo el bloque debajo)

```
Necesito que corras un set de sensibilidades sobre mi modelo financiero usando este kit.

## Mi info

- Archivo del modelo: [NOMBRE_DEL_ARCHIVO.xlsm]
  (está en esta carpeta KIT_SENSIBILIDADES_CONECTAR)
- Nombre del activo / proyecto: [NOMBRE]
- Fecha de valoración: [FECHA, ej: Dic 2025]

## Lo que quiero

Las 16 sensibilidades estándar (ver matriz_escenarios.md) — IPC, IPP, IBR, Ke, OPEX y combinados.
[O si querés algo distinto, descríbelo aquí. Ejemplo:
"Solo Ke +/-1%, +/-2.5%, y un shock de tarifa -5%, -10%"]

## Outputs a trackear (qué celdas registrar por escenario)

Por defecto: Equity DDM (`Valoración!F137`) + IRR + CFADS Total + FCFE Total + PR Total + Caja Final + Caja Mín (waterfall en `EEFF IFRS`).

[Si querés agregar más outputs, lista acá. Ejemplos:
- "Trackea también todos los Ingresos (Total + desglose por contraparte)"
- "Trackea Costos Operacionales y Gastos por línea"
- "Solo me importa: Equity WACC F124, Equity DDM F137, Equity PR F148"
- "Series mensuales para CFADS, FCFE, Pagos Restringidos y Saldo Caja"
- "Trackea Deuda año a año: Financing!J65, T65, AE65, AP65"]

## Series mensuales (opcional)

Si querés trackear series mensuales (no solo totales), pedirlas explícitamente. Útil para auditar coherencia de flujos por escenario. Por defecto se incluyen las 4 series cash flow (CFADS, FCFE, PR, Saldo Caja).

## Post-procesamiento (organizado)

Al final, Claude correrá `build_organized.py` para generar un xlsx limpio multi-hoja con:
- Resumen ejecutivo (top movers, ranking por bloque)
- Tablas detalladas por categoría
- Series mensuales por grupo
- Color coding (verde/rojo), audit cells, freeze panes

## Mi compromiso

- Tengo Excel cerrado y permisos para macros activos
- Estoy disponible si me preguntás algo

## Tu compromiso (Claude)

1. Lee primero `INSTRUCCIONES_CLAUDE.md` y `WORKFLOW_TECNICO.md` en esta carpeta.
2. Mapea inputs del modelo (si es el layout estándar de Conectar, no me preguntes nada — usá los rows típicos).
3. Verifica baseline (corre macro sin shocks, lee Equity DDM) y reportame el número.
4. Si el baseline da error o el modelo se ve no estándar, paramá y preguntáme.
5. Confirmá la matriz de escenarios antes de correr (puedo ajustarla).
6. Corre los 16 escenarios con el workflow J-directo del playbook.
7. Audita pares +/- simétricos al final.
8. Guarda los resultados a `MODELO_[NOMBRE]_SENSIBILIDADES.xlsm` en esta carpeta.
9. Entrégame tabla resumen + análisis de coherencia económica.

Importante: NO me preguntes cosas que están resueltas en `WORKFLOW_TECNICO.md` (workflow J vs K, retry COM, manejo de errores, etc.). SOLO preguntame cosas específicas a mi modelo que no podés inferir.

Comienza.
```

---

## Variantes del prompt según situación

### Si querés que Claude pregunte primero antes de correr

Agrega al final del prompt:

```
Antes de empezar a correr, mostrame el mapeo de inputs que pensás usar (qué rows en Inputs_C corresponden a IPC, Rf, OPEX, etc.) y la matriz de escenarios. Te confirmo y arrancás.
```

### Si querés solo unos pocos escenarios

Reemplaza la sección "Lo que quiero":

```
## Lo que quiero

Solo estos escenarios:
1. Ke +1%, +2.5%
2. Ke -1%, -2.5%
3. Tarifa -5%, -10%
4. Plazo contrato +5 años
```

### Si tenés un shock NO estándar

```
## Lo que quiero

Las 16 estándar + estos custom:
- "Tarifa Otrosi #5 -10%": multiplica Inputs_C!J104 por 0.90
- "Sin extensión de contrato": pone Inputs_C!J45 = 0
```

### Si NO querés que toque la macro (solo lectura)

```
## Lo que quiero

Solo mapeame los inputs y outputs del modelo (sin correr nada). Quiero saber dónde están las celdas para correr yo manualmente después.
```

---

## Ejemplo lleno

```
Necesito que corras un set de sensibilidades sobre mi modelo financiero usando este kit.

## Mi info

- Archivo del modelo: 20260429_Valoracion_ACTIVO_VF1.xlsm
  (está en esta carpeta KIT_SENSIBILIDADES_CONECTAR)
- Nombre del activo / proyecto: Activo de transmisión eléctrica (Colombia)
- Fecha de valoración: Dic 2025

## Lo que quiero

Las 16 sensibilidades estándar (ver matriz_escenarios.md).

## Mi compromiso

- Tengo Excel cerrado y permisos para macros activos
- Estoy disponible si me preguntás algo

## Tu compromiso (Claude)

[...mismo bloque de antes...]

Comienza.
```
