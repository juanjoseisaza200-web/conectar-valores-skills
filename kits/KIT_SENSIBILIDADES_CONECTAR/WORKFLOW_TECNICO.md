# Workflow Sensibilidades — Modelos Project Finance Conectar Valores

Este documento describe el proceso completo y reproducible para correr escenarios de sensibilidad sobre modelos Excel con macro de cierre de circulares. Validado con modelo PEL Mayo 2025 VF1.

---

## 1. Pre-requisitos del modelo

El modelo Excel debe tener:

- **Hoja `Inputs_C`** con estructura columnar de escenarios (cols I-X o similar):
  - Fila 5: encabezados de escenarios (J5 = referencia a J7)
  - Fila 6: `F6` = "Input activo" (selector de escenario activo, texto)
  - Fila 7: `J7` = nombre del escenario (caso base) — debe coincidir con F6
  - `H5 = MATCH(F6, $I$5:$X$5, 0)` (índice columna activa)
  - Cada input usa: `F{r} = INDEX(I{r}:X{r}, $H$5)` para traer valor del escenario activo
  - Columna J = caso base, K-X disponibles como backup/escenarios pre-construidos

- **Hoja `Inputs_S`** con time series y selectores tipo CHOOSE para curvas (IBR, etc.)
  - Curvas alternativas en filas consecutivas + CHOOSE que selecciona según `Inputs_C!F{escenario}`
  - Filas dedicadas a "sensibilidad" (vacías) que el usuario puede sobrescribir

- **Macro `sensibilidad`** (botón en Inputs_C) que:
  - Itera convergencia de circulares (taxes, dividends, DSCR, cash sweep)
  - Hace copy-paste de un rango en hoja Outputs (cierra el modelo)
  - NO recalcula perfiles de amortización (los mantiene fijos)

- **Hoja Resultados** (creada por usuario) para escribir tabla de salidas

---

## 2. Regla de oro

> **Modificar columna J directamente** (no K). Mantener F6 y J7 sincronizados (mismo string).

### Por qué

El modelo tiene celdas downstream que referencian **directamente** `Inputs_C!$J{row}` (no via F-column INDEX). Si modifico K y cambio F6, el modelo queda en estado mixto: parte cells leen K (shocked), parte leen J (base). Resultado: offset sistemático ~7% en equity sin sentido económico, pares +/- ambos del mismo signo.

Al modificar J directamente con K como backup, **todas** las referencias downstream — sean via F-column o directas — apuntan al mismo set de inputs shocked. Resultados con simetría perfecta en pares +/-.

---

## 3. Workflow paso a paso

### Setup (1 vez)

1. Cerrar Excel si está abierto
2. Copiar modelo limpio a carpeta de trabajo (`MODELO_BASE.xlsm` u otro nombre)
3. Verificar baseline: abrir, leer `Valoración!F137` (o output equivalente)
4. Cerrar archivo

### Por cada escenario

```
1. Backup J → K (snapshot de baseline) - solo 1 vez al inicio del set
   $ws_in.Range("J5:J915").Copy($ws_in.Range("K5:K915"))

2. Restaurar J desde K (filas 8+ para no afectar nombre)
   $ws_in.Range("K8:K915").Copy($ws_in.Range("J8:J915"))
   $ws_in.Range("J7").Value = BASE_NAME
   $ws_in.Range("F6").Value = BASE_NAME
   $ws_in.Range("J812").Value = 4  # base IBR escenario
   $ws_s.Range("AA151:BN151").ClearContents()  # limpia IBR sensibilidad

3. Aplicar shocks DIRECTAMENTE en J:
   - IPC: J722:J734 += delta
   - IPP factor: J737 += delta
   - Rf (Ke DDM): J909 += delta
   - OPEX: J347, J349, J382-384, J444 *= factor
   - IBR: escribir R151 = R150 + delta (paste-special values), J812 = 5

4. Cambiar nombre escenario:
   $ws_in.Range("J7").Value = "IBR -1%"
   $ws_in.Range("F6").Value = "IBR -1%"   # DEBEN COINCIDIR

5. Calcular + correr macro:
   $excel.Calculation = -4105   # Automatic
   $excel.CalculateFull()
   $excel.Run("sensibilidad")

6. Leer outputs:
   $eq = $ws_val.Range("F137").Value2
   $nd = $ws_val.Range("F123").Value2
   $ev = $eq + $nd

7. Escribir fila en hoja Resultados
```

### Final

```
1. Restaurar J desde K
2. F6 y J7 = baseline
3. Limpiar K column y R151
4. Correr macro una última vez para dejar modelo en estado consistente
5. Guardar archivo (con NUEVO nombre para evitar OneDrive sync conflicts)
```

---

## 4. Trampas COM PowerShell (críticas)

### TypeCast `Double → String`
Las celdas con NumberFormat heredado del template rechazan asignación numérica.

**Solución:** helper `WriteCell` que setea `cell.NumberFormat = "General"` y usa `.Value` (variant), no `.Value2`.

```powershell
function WriteCell {
    param($ws, $row, $col, $value)
    for ($try = 0; $try -lt 12; $try++) {
        try {
            $cell = $ws.Cells.Item($row, $col)
            try { $cell.Validation.Delete() | Out-Null } catch {}
            $cell.NumberFormat = "General"
            if ($value -is [double] -or $value -is [int]) {
                $cell.Value = [double]$value
            } else {
                $cell.Value = [string]$value
            }
            return
        } catch {
            if ($_.Exception.Message -match "RPC|rechazada|busy|0x80010001") {
                Start-Sleep -Milliseconds (250 * ($try + 1)); continue
            }
            throw
        }
    }
}
```

### RPC_E_CALL_REJECTED (0x80010001)
Excel ocupado durante calc/macro. Solución: wrap todos los reads/writes COM en retry con backoff.

```powershell
function Invoke-COMRetry {
    param([scriptblock]$Action, [int]$MaxRetries = 12, [int]$DelayMs = 250)
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try { return & $Action }
        catch {
            if ($_.Exception.Message -match "0x8001010A|0x80010001|RPC_E|rechazada|busy|0x800A9C68") {
                Start-Sleep -Milliseconds ($DelayMs * ($i + 1)); continue
            }
            throw
        }
    }
}
```

### Funciones con verbo Set-/Get-
Confunden parser PowerShell. Usar `DoCalc`, `RunMacro`, `WriteCell` en lugar de `Set-CalcAuto`, `Run-Macro`.

### `wb.Save()` corrompe el modelo
Si las circulares no convergen al guardar, el archivo queda con `F137 = #N/A` (-2146826246) y pierde ~1MB. **Guardar SOLO al final**, después de restaurar baseline y correr macro completa.

### OneDrive sync conflicts
Si `MODELO_BASE.xlsm` está siendo trabajado por OneDrive, escribir a un nombre nuevo (ej. `MODELO_<empresa>_SENSIBILIDADES.xlsm`) para evitar pisado.

### Python Windows Store
Tiene problemas con COM. Usar PowerShell directo o python sistema (`C:\Users\<user>\AppData\Local\Python\pythoncore-3.14-64\python.exe`).

---

## 5. Auditoría de coherencia económica

Por cada par +/-, validar:

| Check | Esperado |
|---|---|
| **Simetría** | `\|Δ%(+)\| ≈ \|Δ%(-)\|` con leve convexidad para Ke |
| **Direcciones** | IPP↑→Eq↑, IBR↑→Eq↓, Ke↑→Eq↓ (DDM), OPEX↑→Eq↓ |
| **Fisher** | shocks IPC+IBR juntos en misma dirección → impacto cercano a 0 |
| **Ranking magnitudes** | Ke > IBR > OPEX > IPP > IPC Fisher (para DDM) |

**Si algún par muestra ambos signos negativos (o positivos) cuando debería haber simetría, el workflow está mal.** Probable causa: columna K activa cuando debería ser J, o F6 ≠ J7.

---

## 6. Estructura de inputs típicos (modelo PEL)

| Input | Celda Inputs_C | Default | Notas |
|---|---|---|---|
| IPC Colombia | J720:J734 | hist + 5%→3% LP | shock desde primer año proyectado |
| Factor IPC→IPP | J737 | 1.1503 | IPP = IPC × factor |
| Spread Tramos deuda | J612:J615 | 4.10%, 3.90%, 3.90%, 3.50% | sumar a curva base |
| Curva DTF anual | J816:J821 | 5.65%-10.17% | proyección FDN |
| Escenario IBR | J812 | 4 (Caso Base) | selector de CHOOSE |
| Rf | J909 | 4.30% | base CAPM |
| Beta unlevered | J910 | 0.5 | Damodaran |
| ERP | J911 | 5.50% | |
| CRP Colombia | J912 | 2.00% | |
| Size premium | J913 | 1.50% | |
| OPEX personal | J347 | 1,585 COP MM | shock x factor |
| OPEX mant.+rep. | J349 | 7,712 COP MM | shock x factor |
| OPEX Statcom | J382 | 105 COP MM/mes | shock x factor |
| OPEX C.Control | J383 | 73.5 COP MM/mes | shock x factor |
| OPEX Corocora | J384 | 158 COP MM/mes | shock x factor |
| Mant. mayores | J444 | 1,771.5 COP MM 2025 | shock x factor |

### IBR Sensibilidad (Inputs_S)
- R150 = "IBR Caso Base" (cols ~AA-BN, valores hardcoded)
- R151 = "IBR Sensibilidad" (vacía, escribir aquí)
- R152 = "IBR Sensibilidad" (vacía, alternativa)
- R154 = `=CHOOSE($F$145, J147, J148, J149, J150, J151, J152)` (selector)
- F145 = `Inputs_C!F812` (escenario IBR activo)

Para shock IBR: `R151 = R150 + delta` (paste-special values), J812 = 5.

---

## 7. Outputs DDM

| Métrica | Celda |
|---|---|
| Equity Value (DDM) | `Valoración!F137` |
| Deuda Neta (Dic 2025) | `Valoración!F123` |
| EV implícito (DDM) | F137 + F123 (calcular) |
| Ke anual periodo a periodo | `Valoración!J86:KP86` |

Baseline esperado PEL: Equity DDM = 124,334 / EV = 257,581 COP MM.

---

## 8. Matriz de sensibilidades estándar (16 escenarios)

| # | Escenario | Shock primario | Acompañantes (Fisher) |
|---|---|---|---|
| 1 | IPC +1% Fisher | J722:J734 += 0.01 | IBR +1% (R151+0.01, J812=5) |
| 2 | IPC -1% Fisher | J722:J734 -= 0.01 | IBR -1% |
| 3 | Spread IPP +1% | J737 += 0.01 | — |
| 4 | Spread IPP -1% | J737 -= 0.01 | — |
| 5 | IBR +1% puro | R151 = R150+0.01, J812=5 | — |
| 6 | IBR -1% puro | R151 = R150-0.01, J812=5 | — |
| 7 | IBR +2.5% | R151 = R150+0.025 | — |
| 8 | IBR -2.5% | R151 = R150-0.025 | — |
| 9 | Ke +1% | J909 += 0.01 | — |
| 10 | Ke -1% | J909 -= 0.01 | — |
| 11 | Ke +2.5% | J909 += 0.025 | — |
| 12 | Ke -2.5% | J909 -= 0.025 | — |
| 13 | OPEX +10% | J347,J349,J382-384,J444 *= 1.10 | — |
| 14 | OPEX -10% | OPEX *= 0.90 | — |
| 15 | Downside combinado | IPC+1%, IBR+1%, Ke+1%, OPEX+10% | Fisher coherente |
| 16 | Upside combinado | IPC-0.5%, IBR-0.5%, Ke-1%, OPEX-10% | Fisher coherente |

---

## 9. Resultados validados (PEL Mayo 2025 VF1)

| # | Escenario | Equity DDM | Δ% Eq |
|---|---|---|---|
| 0 | BASELINE | 124,334 | 0% |
| 1 | IPC +1% Fisher | 123,790 | -0.44% |
| 2 | IPC -1% Fisher | 124,486 | +0.12% |
| 3 | Spread IPP +1% | 125,099 | +0.62% |
| 4 | Spread IPP -1% | 123,571 | -0.61% |
| 5 | IBR +1% | 122,219 | -1.70% |
| 6 | IBR -1% | 126,464 | +1.71% |
| 7 | IBR +2.5% | 119,072 | -4.23% |
| 8 | IBR -2.5% | 129,686 | +4.30% |
| 9 | Ke +1% | 116,471 | -6.32% |
| 10 | Ke -1% | 132,795 | +6.81% |
| 11 | Ke +2.5% | 105,689 | -15.00% |
| 12 | Ke -2.5% | 146,722 | +18.01% |
| 13 | OPEX +10% | 122,546 | -1.44% |
| 14 | OPEX -10% | 126,123 | +1.44% |
| 15 | Downside combinado | 114,048 | -8.27% |
| 16 | Upside combinado | 134,692 | +8.33% |

**Auditoría pares simétricos: PASS.** Todos los pares muestran simetría correcta y direcciones económicas válidas.

---

## 10. Archivos del proceso

| Archivo | Uso |
|---|---|
| `MODELO_BASE (1).xlsm` | Modelo limpio inicial (con R151/R152 IBR Sens + Hoja Resultados) |
| `MODELO_PEL_SENSIBILIDADES.xlsm` | Modelo con resultados grabados |
| `run_sensibilidades_template.ps1` | Script template adaptable |
| `WORKFLOW_SENSIBILIDADES.md` | Este documento |
| `prompt_template.md` | Prompt template para invocar el proceso |
