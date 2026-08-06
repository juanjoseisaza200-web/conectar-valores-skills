# INSTRUCCIONES PARA CLAUDE — Lee esto PRIMERO

Sos Claude trabajando con un modelo financiero project finance de Conectar Valores S.A.S. El usuario te pasó este kit `KIT_SENSIBILIDADES_CONECTAR/` con todo lo necesario para correr sensibilidades.

## Tu mandato

Correr 16 escenarios de sensibilidad (o los que el usuario pida) sobre un modelo Excel `.xlsm`, con el workflow J-directo validado, y entregar resultados auditados en una hoja del modelo.

## Lo que YA SABÉS (no preguntes al usuario)

### Estructura típica del modelo Conectar Valores
Los modelos siguen un patrón estándar:
- **Hoja `Inputs_C`**: estructura columnar de escenarios (cols I-X o más)
  - F6 = "Input activo" (selector, texto)
  - J7 = nombre del escenario activo (caso base)
  - J5 = `=J7` (header dinámico)
  - H5 = `MATCH(F6, $I$5:$X$5, 0)` (índice de columna activa)
  - F-column: `INDEX(I{r}:X{r}, $H$5)` para traer valor del escenario activo
- **Hoja `Inputs_S`**: time series con CHOOSE selectors para curvas IBR
  - R150 = "IBR Caso Base" (cols ~AA-BN)
  - R151, R152 = "IBR Sensibilidad" (vacías para shocks)
  - R154 = `CHOOSE($F$145, J147, J148, J149, J150, J151, J152)`
  - F145 = `Inputs_C!F812`
- **Macro `sensibilidad`**: itera convergencia + copy-paste Outputs + cierra
- **Hoja `Valoración`**: F137 = Equity DDM, F123 = Deuda Neta

### Workflow correcto (CRITICAL)
1. **Modificar columna J directamente**, no K. K es backup del baseline.
2. **F6 y J7 deben coincidir** con el nombre del escenario.
3. Copy J→K una vez al inicio, restaurar K→J entre escenarios (filas 8+).
4. Macro `sensibilidad` maneja todo internamente — no necesita CalculateFull antes.
5. NO `wb.Save()` durante el loop — solo al final tras restaurar baseline.
6. Wrap todo COM call en retry (RPC busy + 0x800A9C68 macro errors).

### Inputs estándar (rows en Inputs_C, columna J)
| Variable | Row | Notes |
|---|---|---|
| IPC Colombia proyectado | 720-734 | shock desde primer año proy (típicamente row 722) |
| Factor IPC→IPP | 737 | IPP = IPC × factor |
| Spreads de deuda | 612-615 | tramos 1-4 |
| Curva DTF | 816-821 | proyecciones FDN |
| Selector escenario IBR | 812 | default 4 (Caso Base) |
| Rf | 909 | base CAPM |
| Beta | 910 | Damodaran |
| ERP | 911 | |
| CRP | 912 | |
| Size premium | 913 | |
| OPEX personal | 347 | |
| OPEX mant. y rep. | 349 | |
| OPEX Statcom | 382 | mensual |
| OPEX C.Control | 383 | mensual |
| OPEX Corocora | 384 | mensual |
| Mant. mayores | 444 | |

**Si el modelo es ESTÁNDAR (PEL/Conectar)**, usa estos rows tal cual. **No preguntes**.

### Outputs estándar (por defecto)
- `Valoración!F137` = Equity Value (DDM)
- `Valoración!F123` = Deuda Neta
- EV (DDM) = F137 + F123 (calcular)

### Outputs CONFIGURABLES (`$OUTPUTS` en el script)

El script soporta cualquier número de outputs (cualquier celda de cualquier hoja). Si el usuario pidió outputs custom, edita `$OUTPUTS` en `run_sensibilidades.ps1`:

```powershell
$OUTPUTS = @(
    @{label="Equity DDM";   sheet="Valoración"; cell="F137"; type="value"},
    @{label="Deuda Neta";   sheet="Valoración"; cell="F123"; type="value"},
    @{label="DSCR min";     sheet="Financing";  cell="F250"; type="value"},
    @{label="FCF total";    sheet="Valoración"; cell="F54";  type="value"}
)
```

El script genera 2 columnas por output (valor + Δ%) en la hoja Resultados automáticamente.

**Outputs típicos por categoría** (cuando el usuario pida algo y no especifique celda):

| Categoría | Output | Celda típica |
|---|---|---|
| Valoración | Equity DDM | Valoración!F137 |
| Valoración | Equity WACC | Valoración!F124 |
| Valoración | Equity PR | Valoración!F148 |
| Valoración | EV WACC | Valoración!F117 |
| Valoración | Promedio 3 métodos | Valoración!F159 |
| Deuda | Deuda Neta | Valoración!F123 |
| Deuda | Saldo deuda final 2037 | Financing!KP65 (último periodo) |
| Deuda | Servicio deuda total | Financing!F<row> (sumarizado) |
| Deuda | DSCR mínimo | Financing buscar "DSCR" |
| Flujo | FCF total | Valoración!F54 (suma) |
| Flujo | FCFE total | Valoración!F131 (suma) |
| Flujo | Pagos restringidos total | Valoración!F148 |
| Operacional | EBITDA año X | EEFF IFRS columna del año |
| Operacional | Ingresos año X | Ingresos columna del año |
| Equity | Total dividendos | Equity!F<row> |
| Tax | Total impuestos | Tx!F<row> |

Si el usuario pide algo que no podés ubicar, **preguntale la celda exacta** o haz `grep` en columna E de la hoja relevante.

### Trampas COM PowerShell ya resueltas
Lee `WORKFLOW_TECNICO.md` para detalle. Resumen:
- TypeCast Double→String: usar helper `WriteCell` con NumberFormat=General + .Value (variant)
- RPC busy 0x80010001: wrap en `Invoke-COMRetry` con backoff
- **READS también** (no solo writes): usar `ReadCellRetry` para `$ws.Cells.Item().Value2`. Síntoma: falla en escenarios 5-15 al leer input.
- Macro error 0x800A9C68: también incluir en retry pattern
- Funciones con verbo Set-/Get-: usar `DoCalc`, `WriteCell` (sin guión)
- OneDrive sync conflicts: usar nombre de salida nuevo (no sobrescribir el original)
- **CalculateFull POST-macro + Sleep 500ms** antes de leer escalares. Sin esto, en escenarios con flags estructurales (ej. extensión contractual) los reads vuelven 0 stale.
- **Save antes de format step**. Si format falla (ej. cols >Z mal manejadas), datos preservados.
- **IRR con guess explícito**: usar `IRR(range, 0.01)` no `IRR(range)` solo. Sin guess Excel a veces no converge y devuelve #N/A.
- **Audit caja relativo a baseline**: si baseline tiene caja negativa, no usar threshold absoluto `<-50`. Usar `< (base_caja_min - 1000)` (empeora >1k vs base).

## Lo que SÍ debés preguntar al usuario (solo si aplica)

1. **Nombre exacto del archivo** si hay varios .xlsm en la carpeta
2. **Nombre de la macro** si NO se llama `sensibilidad` (algunos modelos antiguos pueden tener variantes)
3. **Confirmación de matriz de escenarios** si el usuario pidió "los estándar" pero querés validar (mostrar la lista y esperar OK)
4. **Shocks específicos del activo** si el usuario menciona algo no estándar (ej: tarifa, capacidad, plazos)
5. **Si el modelo NO es estándar** (no tiene la estructura PEL típica): ahí sí mapear con grep y validar con usuario

## Lo que NO debés preguntar (lo decidís solo)

- Workflow J vs K (siempre J directo)
- Si correr macro o no (siempre correr `sensibilidad`)
- Si hacer auditoría de pares (siempre)
- Si guardar a archivo nuevo (siempre, para evitar OneDrive conflicts)
- Detalles de retry COM
- Detalles de NumberFormat
- Si verificar baseline antes de empezar (siempre verificar contra el valor que el usuario o memoria indique)

## Workflow ejecutivo (para vos)

```
1. Leer WORKFLOW_TECNICO.md (si aún no lo has leído en esta sesión)
2. Identificar archivo del modelo en la carpeta del kit
3. Verificar estructura: abrir con openpyxl, listar hojas, confirmar Inputs_C/Inputs_S/Valoración existen
4. Mapear inputs: si labels matching estándar Conectar (grep "IPC", "Factor IPC", "Tasa Libre", "OPEX"), usar rows estándar. Si NO match, preguntar al usuario.
5. Verificar baseline: abrir, F6=baseline, run macro, leer F137. Reportar al usuario.
   - Si F137 = error o muy diferente al esperado: PARAR, reportar.
6. Confirmar matriz con usuario (mostrar 16 escenarios, esperar GO).
7. Adaptar `run_sensibilidades.ps1`: editar $CFG con paths/nombres específicos del modelo.
8. Ejecutar el script (background con monitor).
9. Trackear progreso por escenario, reportar cada uno.
10. Al final: auditoría de pares, ranking de sensibilidades, archivo guardado.
```

## Qué entregar

1. **Archivo `.xlsm`** con resultados en hoja Resultados (filas 12-29 típicas)
2. **Tabla resumen en chat** con 17 filas (BASELINE + 16) y deltas
3. **Análisis de pares simétricos**: cada par +/- debe mostrar simetría aprox. correcta
4. **Ranking de sensibilidad** (Ke típicamente domina, IBR segundo, etc.)
5. **Comentarios económicos**: Fisher debe cumplirse, direcciones correctas

## Señales de alarma (PARA inmediatamente)

- Baseline F137 = error (-2146826246) → modelo corrupto
- Pares +/- ambos negativos sin sentido → workflow incorrecto (probablemente K activa)
- Macro falla 12 veces seguidas con 0x800A9C68 → problema estructural
- Equity Value < 0 o > 10x baseline → bug en shock

## Si el usuario quiere otros escenarios

- Escenarios CUSTOM (no en la matriz estándar): pedir qué celda y qué shock. Implementar como modificación J directa con misma mecánica.
- Escenarios de **operación** (capacidad MW, plazos contrato): aclarar con usuario porque pueden romper otras lógicas del modelo. Suelen requerir cambio de FLAGS (J32-J38, J45) más que valores.
- Escenarios sobre OTROS activos (Geopark, DELSUR, TRANSELCO en PEL): mapear OPEX rows distintos. Ver `WORKFLOW_TECNICO.md`.

## Recordatorio final

El usuario te dio un PROMPT_INICIAL.md con su info. Si NO menciona algo específico, asumí estructura estándar. **No le hagas perder tiempo con preguntas básicas.** Solo preguntale lo que es realmente necesario para vos avanzar.

Si la corrida sale exitosa, **guarda en memoria persistente** (si tu Claude soporta auto-memory) los resultados de esta corrida. Eso ayuda a futuras conversaciones del mismo proyecto.
