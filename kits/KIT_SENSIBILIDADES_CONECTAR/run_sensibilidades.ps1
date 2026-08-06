# =============================================================================
# RUN SENSIBILIDADES - Template adaptable a cualquier modelo PF Conectar
# =============================================================================
# Workflow validado: modificar J directo, K como backup, F6 y J7 coinciden.
# Macro `sensibilidad` cierra circulares + copy-paste Outputs.
#
# Adaptar las VARIABLES DE CONFIGURACION a tu modelo y ejecutar.
# =============================================================================

$ErrorActionPreference = "Stop"
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# =============================================================================
# 1. VARIABLES DE CONFIGURACION (AJUSTAR PER MODELO)
# =============================================================================

# NOTA: Por defecto el script usa rutas relativas a la carpeta del kit.
# Reemplaza $InputFile con el nombre exacto de tu modelo en esta carpeta.
$InputFile = "MODELO_INPUT.xlsm"   # <<< CAMBIAR al nombre real
$OutputFile = "MODELO_RESULTADOS_SENSIBILIDADES.xlsm"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$CFG = @{
    # Archivos (resueltos a partir de la carpeta del kit)
    PathSrc      = Join-Path $ScriptDir $InputFile
    PathDst      = Join-Path $ScriptDir $OutputFile

    # Hojas
    SheetInputs  = "Inputs_C"
    SheetTimeSer = "Inputs_S"
    SheetVal     = "Valoración"
    SheetResult  = "Reusltado Sensibilidades"

    # Selectores
    BaseName     = "Caso Base Tramo 1 y 2 "  # incluir trailing space si aplica
    MacroName    = "sensibilidad"

    # Filas inputs (ajustar)
    IpcRowStart  = 722  # primer año proyectado
    IpcRowEnd    = 734  # último año
    IppFactorRow = 737
    RfRow        = 909
    IbrScenRow   = 812  # selector escenario IBR (default 4=base)
    OpexRows     = @(347, 349, 382, 383, 384, 444)  # OPEX inputs (layout estandar)

    # IBR Sensibilidad (Inputs_S)
    IbrBaseRow   = 150  # IBR Caso Base (hardcoded)
    IbrSensRow   = 151  # IBR Sensibilidad (escribir aquí)
    IbrSensRange = "AA151:BN151"  # rango cols donde shock se aplica
    IbrBaseRange = "AA150:BN150"
    IbrSensCode  = 5    # valor de IbrScenRow para activar IbrSensRow

    # Resultados (zona de escritura)
    ResStartRow  = 12
    ResHeaderRow = 12
    ResDataRow   = 13  # primera fila de datos (BASELINE)
}

# =============================================================================
# 1.b OUTPUTS A TRACKEAR (CONFIGURABLE)
# =============================================================================
# Lista de outputs. Cada uno tiene:
#   label: nombre legible (cabecera de columna en Resultados)
#   sheet: nombre de la hoja
#   cell:  celda (notación A1, ej "F137")
#   type:  "value" o "pct" (para formato de salida)
# Para cada output, el script mostrará: [valor], [delta vs base], [delta % vs base]
#
# EJEMPLOS:
#   Equity Value DDM:    label="Equity DDM",    sheet="Valoración",   cell="F137"
#   Deuda Neta:          label="Deuda Neta",    sheet="Valoración",   cell="F123"
#   FCF año 5:           label="FCF Y5",        sheet="Valoración",   cell="N54"
#   DSCR mínimo:         label="DSCR min",      sheet="Financing",    cell="F250"
#   Saldo deuda final:   label="Deuda 2037",    sheet="Financing",    cell="KP65"
#   Total dividendos:    label="Div Total",     sheet="Equity",       cell="F30"
#   TIR Equity:          label="TIR Eq",        sheet="Valoración",   cell="F165"
#   EBITDA año 1:        label="EBITDA Y1",     sheet="EEFF IFRS",    cell="J50"

# OUTPUTS soporta DOS modos:
# 1) Celda única:  @{label=...; sheet=...; cell="F137"}
# 2) Rango con agregación: @{label=...; sheet=...; range="J157:KP157"; agg="sum|min|last"}
#    Útil para trackear TOTALES de series mensuales (ingresos, opex, FCFE total, etc.)
$OUTPUTS = @(
    # Valoración (celdas)
    @{label="Equity DDM";  sheet="Valoración"; cell="F137"},
    @{label="IRR Eq";      sheet="Valoración"; cell="F168"}, # se computa en pre-flight
    # Cash flow waterfall (rangos sumarizados)
    @{label="CFADS Total";  sheet="EEFF IFRS"; range="J157:KP157"; agg="sum"},
    @{label="FCFE Total";   sheet="EEFF IFRS"; range="J198:KP198"; agg="sum"},
    @{label="PR Total";     sheet="EEFF IFRS"; range="J200:KP200"; agg="sum"},
    @{label="Caja Final";   sheet="EEFF IFRS"; range="J204:KP204"; agg="last"},
    @{label="Caja Min";     sheet="EEFF IFRS"; range="J204:KP204"; agg="min"}
    # Agregar más outputs aquí. Ejemplos:
    # @{label="Ingresos Tot"; sheet="Ingresos"; range="J217:KP217"; agg="sum"},
    # @{label="Opex Total";   sheet="Opex";     range="J248:KP248"; agg="sum"},
    # @{label="Eq WACC";      sheet="Valoración"; cell="F124"}
)

# =============================================================================
# 1.c SERIES MENSUALES A TRACKEAR (opcional)
# =============================================================================
# Por cada serie, el script copia el rango J:KP de la fila origen a la hoja
# Resultados como bloque mensual (filas 56 escenarios x ~290 cols mensuales).
# Si no querés series mensuales, dejá $SERIES = @().
$SERIES = @(
    @{label="CFADS";              sheet="EEFF IFRS"; row=157},
    @{label="FCFE (Disp Acc)";    sheet="EEFF IFRS"; row=198},
    @{label="Pagos Restringidos"; sheet="EEFF IFRS"; row=200},
    @{label="Saldo Caja";         sheet="EEFF IFRS"; row=204}
    # Agregar más series aquí. Ejemplos:
    # @{label="Ingresos Total";     sheet="Ingresos"; row=217},
    # @{label="Opex Total";         sheet="Opex";     row=248},
    # @{label="Costos Op";          sheet="Opex";     row=118}
)

# =============================================================================
# 2. MATRIZ DE ESCENARIOS (16 estándar - editable)
# =============================================================================

$SCENARIOS = @(
    @{n="IPC +1% (Fisher)";    ipc=0.01;   ipp=0;      ibr=0.01;   rf=0;       opex=1.0;  notes="+1pp IPC y +1pp IBR (Fisher)"},
    @{n="IPC -1% (Fisher)";    ipc=-0.01;  ipp=0;      ibr=-0.01;  rf=0;       opex=1.0;  notes="-1pp IPC y -1pp IBR"},
    @{n="Spread IPP +1%";      ipc=0;      ipp=0.01;   ibr=0;      rf=0;       opex=1.0;  notes="+1pp factor IPC->IPP"},
    @{n="Spread IPP -1%";      ipc=0;      ipp=-0.01;  ibr=0;      rf=0;       opex=1.0;  notes="-1pp factor IPC->IPP"},
    @{n="IBR +1%";             ipc=0;      ipp=0;      ibr=0.01;   rf=0;       opex=1.0;  notes="+1pp IBR puro"},
    @{n="IBR -1%";             ipc=0;      ipp=0;      ibr=-0.01;  rf=0;       opex=1.0;  notes="-1pp IBR puro"},
    @{n="IBR +2.5%";           ipc=0;      ipp=0;      ibr=0.025;  rf=0;       opex=1.0;  notes="Stress crediticio"},
    @{n="IBR -2.5%";           ipc=0;      ipp=0;      ibr=-0.025; rf=0;       opex=1.0;  notes="Upside crediticio"},
    @{n="Ke +1% (DDM)";        ipc=0;      ipp=0;      ibr=0;      rf=0.01;    opex=1.0;  notes="+1pp Ke"},
    @{n="Ke -1% (DDM)";        ipc=0;      ipp=0;      ibr=0;      rf=-0.01;   opex=1.0;  notes="-1pp Ke"},
    @{n="Ke +2.5% (DDM)";      ipc=0;      ipp=0;      ibr=0;      rf=0.025;   opex=1.0;  notes="Stress descuento"},
    @{n="Ke -2.5% (DDM)";      ipc=0;      ipp=0;      ibr=0;      rf=-0.025;  opex=1.0;  notes="Upside descuento"},
    @{n="OPEX +10%";           ipc=0;      ipp=0;      ibr=0;      rf=0;       opex=1.10; notes="+10% OPEX"},
    @{n="OPEX -10%";           ipc=0;      ipp=0;      ibr=0;      rf=0;       opex=0.90; notes="-10% OPEX"},
    @{n="Downside combinado";  ipc=0.01;   ipp=0;      ibr=0.01;   rf=0.01;    opex=1.10; notes="Stress integral Fisher"},
    @{n="Upside combinado";    ipc=-0.005; ipp=0;      ibr=-0.005; rf=-0.01;   opex=0.90; notes="Best case Fisher"}
)

# =============================================================================
# 3. HELPERS COM
# =============================================================================

function Invoke-COMRetry {
    param([scriptblock]$Action, [int]$MaxRetries = 12, [int]$DelayMs = 250)
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try { return & $Action }
        catch {
            $msg = $_.Exception.Message
            if ($msg -match "0x8001010A|0x80010001|RPC_E|rechazada|busy|0x800A9C68") {
                Start-Sleep -Milliseconds ($DelayMs * ($i + 1)); continue
            }
            throw
        }
    }
    throw "COM call failed after $MaxRetries retries"
}

function ReadCellRetry($ws, $row, $col) {
    for ($t = 0; $t -lt 12; $t++) {
        try { return $ws.Cells.Item($row, $col).Value2 }
        catch {
            if ($_.Exception.Message -match "RPC|rechazada|busy|0x80010001") {
                Start-Sleep -Milliseconds (250 * ($t + 1)); continue
            }
            throw
        }
    }
}

function WriteCell {
    param($ws, $row, $col, $value)
    for ($try = 0; $try -lt 12; $try++) {
        try {
            $cell = $ws.Cells.Item($row, $col)
            try { $cell.Validation.Delete() | Out-Null } catch {}
            $cell.NumberFormat = "General"
            if ($value -is [double] -or $value -is [int] -or $value -is [int64] -or $value -is [decimal]) {
                $cell.Value = [double]$value
            } else {
                $cell.Value = [string]$value
            }
            return
        } catch {
            if ($_.Exception.Message -match "0x80010001|RPC_E|rechazada|busy") {
                Start-Sleep -Milliseconds (250 * ($try + 1)); continue
            }
            throw
        }
    }
}

# Helpers para rangos (sum, min, last non-zero)
function SumRange($ws, $rangeStr) {
    $arr = $ws.Range($rangeStr).Value2
    if ($arr -eq $null) { return 0 }
    $total = 0.0
    if ($arr -is [System.Array]) {
        for ($i = 1; $i -le $arr.GetLength(1); $i++) {
            $v = $arr[1, $i]
            if ($v -ne $null -and $v -is [double]) { $total += $v }
        }
    } else { $total = [double]$arr }
    return $total
}
function MinRange($ws, $rangeStr) {
    $arr = $ws.Range($rangeStr).Value2
    if ($arr -eq $null) { return 0 }
    $min = [double]::MaxValue
    if ($arr -is [System.Array]) {
        for ($i = 1; $i -le $arr.GetLength(1); $i++) {
            $v = $arr[1, $i]
            if ($v -ne $null -and $v -is [double] -and $v -lt $min) { $min = $v }
        }
    } else { $min = [double]$arr }
    if ($min -eq [double]::MaxValue) { return 0 }
    return $min
}
function LastNonZeroRange($ws, $rangeStr) {
    $arr = $ws.Range($rangeStr).Value2
    if ($arr -eq $null) { return 0 }
    $last = 0.0
    if ($arr -is [System.Array]) {
        for ($i = 1; $i -le $arr.GetLength(1); $i++) {
            $v = $arr[1, $i]
            if ($v -ne $null -and $v -is [double] -and $v -ne 0) { $last = $v }
        }
    } else { $last = [double]$arr }
    return $last
}
function ReadOutput($wb, $out) {
    $ws = $wb.Sheets.Item($out.sheet)
    if ($out.cell) {
        $v = $ws.Range($out.cell).Value2
        if ($v -eq $null) { return 0 }
        return [double]$v
    }
    if ($out.range) {
        switch ($out.agg) {
            "sum"  { return SumRange $ws $out.range }
            "min"  { return MinRange $ws $out.range }
            "last" { return LastNonZeroRange $ws $out.range }
        }
    }
    return 0
}
# Conversion column index -> letra (con soporte 2-letter para cols >Z)
function ColLetter($n) {
    if ($n -le 26) { return [string][char]([int][char]'A' + $n - 1) }
    $first = [int](($n - 1) / 26)
    $second = (($n - 1) % 26) + 1
    return ([string][char]([int][char]'A' + $first - 1)) + ([string][char]([int][char]'A' + $second - 1))
}

# =============================================================================
# 4. MAIN
# =============================================================================

Write-Host "Copiando archivo origen a destino..."
Copy-Item $CFG.PathSrc $CFG.PathDst -Force

Write-Host "Abriendo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 1
$excel.EnableEvents = $false
$wb = $excel.Workbooks.Open($CFG.PathDst, $false, $false)
Invoke-COMRetry { $excel.Calculation = -4105 }

$ws_in  = $wb.Sheets.Item($CFG.SheetInputs)
$ws_s   = $wb.Sheets.Item($CFG.SheetTimeSer)
$ws_val = $wb.Sheets.Item($CFG.SheetVal)
$ws_res = $wb.Sheets.Item($CFG.SheetResult)

# Limpieza inicial
Write-Host "Limpieza inicial..."
$ws_in.Range("F6").Value2 = $CFG.BaseName
$ws_in.Range("K5:K915").ClearContents() | Out-Null
$ws_s.Range($CFG.IbrSensRange).ClearContents() | Out-Null
Invoke-COMRetry { $excel.CalculateFull() }

# Backup J -> K
Write-Host "Backup J -> K (snapshot baseline)"
$ws_in.Range("J5:J915").Copy($ws_in.Range("K5:K915")) | Out-Null
$excel.CutCopyMode = $false

# OPCIONAL: pre-flight para IRR (descomenta si querés trackear IRR Equity)
# Crea cash flow stream en Valoración row 167 + IRR formula en F168
# Importante: IRR usa guess 0.01 para evitar non-convergence
# $ws_val.Range("E167").Value = "Cash flow stream para IRR"
# $ws_val.Range("J167:KP167").Formula = "=IF(SUM(`$J`$16:J16)=0, 0, IF(J16=1, -`$F`$137 + J131, J131))"
# $ws_val.Range("E168").Value = "IRR Equity (anualizado)"
# $ws_val.Range("F168").Formula = "=IFERROR((1+IRR(J167:KP167, 0.01))^12-1, NA())"
# Invoke-COMRetry { $excel.CalculateFull() }

# BASELINE - correr macro y leer outputs
Write-Host "`n=== BASELINE ==="
Invoke-COMRetry { $excel.Run($CFG.MacroName) }
Invoke-COMRetry { $excel.CalculateFull() }   # CRITICAL: recalc post-macro evita reads stale (CFADS=0)
Start-Sleep -Milliseconds 500
$base_values = @{}
foreach ($out in $OUTPUTS) {
    $v = ReadOutput $wb $out
    $base_values[$out.label] = $v
    Write-Host ("  {0} ({1}!{2}) = {3:N2}" -f $out.label, $out.sheet, $out.cell, $v)
}

# Setup Resultados header dinamico
# Columnas: # | Escenario | Notas | <out1> | Δ% <out1> | <out2> | Δ% <out2> | ... | Audit
$rng_res = $ws_res.Range("A12:Z80")
$rng_res.Clear() | Out-Null
$rng_res.NumberFormat = "General"

# Cabecera fija
$ws_res.Cells.Item(12, 1).Formula = "#"
$ws_res.Cells.Item(12, 2).Formula = "Escenario"
$ws_res.Cells.Item(12, 3).Formula = "Notas"
# Cabeceras dinamicas: 2 columnas por output (valor + delta %)
$colIdx = 4
$outColMap = @{}  # label -> {valCol, pctCol}
foreach ($out in $OUTPUTS) {
    $ws_res.Cells.Item(12, $colIdx).Formula = $out.label
    $ws_res.Cells.Item(12, $colIdx + 1).Formula = "Δ% " + $out.label
    $outColMap[$out.label] = @{valCol = $colIdx; pctCol = $colIdx + 1}
    $colIdx += 2
}
$auditCol = $colIdx
$ws_res.Cells.Item(12, $auditCol).Formula = "Audit"
$lastColLetter = [char]([int][char]'A' + $auditCol - 1)
if ($auditCol -gt 26) { $lastColLetter = "A" + [char]([int][char]'A' + $auditCol - 27) }
$ws_res.Range("A12:$($lastColLetter)12").Font.Bold = $true

# Fila BASELINE
WriteCell $ws_res 13 1 0
WriteCell $ws_res 13 2 "BASELINE"
WriteCell $ws_res 13 3 "Sin shocks"
foreach ($out in $OUTPUTS) {
    $col = $outColMap[$out.label]
    WriteCell $ws_res 13 $col.valCol $base_values[$out.label]
    WriteCell $ws_res 13 $col.pctCol 0
}
$ws_res.Cells.Item(13, $auditCol).Formula = "BASE"

# LOOP escenarios
$idx = 0
foreach ($s in $SCENARIOS) {
    $idx++
    Write-Host ("`n[{0}/{1}] {2}" -f $idx, $SCENARIOS.Count, $s.n)
    $t0 = Get-Date

    # Restaurar J desde K (filas 8+)
    Invoke-COMRetry { $script:ws_in.Range("K8:K915").Copy($script:ws_in.Range("J8:J915")) | Out-Null }
    $excel.CutCopyMode = $false
    Invoke-COMRetry { $script:ws_s.Range($script:CFG.IbrSensRange).ClearContents() | Out-Null }
    Invoke-COMRetry { $script:ws_in.Range("J7").Value2 = $script:CFG.BaseName }
    Invoke-COMRetry { $script:ws_in.Range("F6").Value2 = $script:CFG.BaseName }
    Invoke-COMRetry { $script:ws_in.Range("J$($script:CFG.IbrScenRow)").Value2 = 4 }

    # Aplicar shocks en J
    if ($s.ipc -ne 0) {
        for ($r = $CFG.IpcRowStart; $r -le $CFG.IpcRowEnd; $r++) {
            $b = ReadCellRetry $ws_in $r 10
            if ($b -ne $null) { WriteCell $ws_in $r 10 ([double]$b + $s.ipc) }
        }
    }
    if ($s.ipp -ne 0) {
        $b = ReadCellRetry $ws_in $CFG.IppFactorRow 10
        WriteCell $ws_in $CFG.IppFactorRow 10 ([double]$b + $s.ipp)
    }
    if ($s.rf -ne 0) {
        $b = ReadCellRetry $ws_in $CFG.RfRow 10
        WriteCell $ws_in $CFG.RfRow 10 ([double]$b + $s.rf)
    }
    if ($s.ibr -ne 0) {
        Invoke-COMRetry { $script:ws_s.Range($script:CFG.IbrSensRange).Formula = "=$($script:CFG.IbrBaseRange.Split(':')[0])+($($s.ibr))" }
        Invoke-COMRetry { $script:excel.CalculateFull() }
        Start-Sleep -Milliseconds 300
        Invoke-COMRetry { $script:ws_s.Range($script:CFG.IbrSensRange).Copy() | Out-Null }
        Invoke-COMRetry { $script:ws_s.Range($script:CFG.IbrSensRange).PasteSpecial(-4163) | Out-Null }
        $excel.CutCopyMode = $false
        WriteCell $ws_in $CFG.IbrScenRow 10 $CFG.IbrSensCode
    }
    if ($s.opex -ne 1.0) {
        foreach ($r in $CFG.OpexRows) {
            $b = ReadCellRetry $ws_in $r 10
            if ($b -ne $null -and $b -is [double]) {
                WriteCell $ws_in $r 10 ([double]$b * $s.opex)
            }
        }
    }

    # Cambiar J7 y F6 al nombre del escenario (DEBEN COINCIDIR)
    Invoke-COMRetry { $script:ws_in.Range("J7").Value2 = $s.n }
    Invoke-COMRetry { $script:ws_in.Range("F6").Value2 = $s.n }
    Start-Sleep -Milliseconds 500

    # Run macro - leer todos los outputs configurados
    try {
        Invoke-COMRetry { $script:excel.CalculateFull() }
        Invoke-COMRetry { $script:excel.Run($script:CFG.MacroName) }
        # CRITICAL: recalc post-macro evita reads stale (síntoma: CFADS=0 random)
        Invoke-COMRetry { $script:excel.CalculateFull() }
        Start-Sleep -Milliseconds 500

        $row = 13 + $idx
        WriteCell $ws_res $row 1 $idx
        WriteCell $ws_res $row 2 $s.n
        WriteCell $ws_res $row 3 $s.notes

        $audit = "OK"
        $primary_pct = 0
        $first = $true
        $log_outputs = ""
        foreach ($out in $OUTPUTS) {
            $v = ReadOutput $wb $out
            $b = $base_values[$out.label]
            $d_pct = if ($b -ne 0) { ($v - $b) / $b } else { 0 }
            $col = $outColMap[$out.label]
            WriteCell $ws_res $row $col.valCol $v
            WriteCell $ws_res $row $col.pctCol $d_pct
            if ($first) { $primary_pct = $d_pct; $first = $false }
            $log_outputs += ("{0}={1:N0} ({2:P2})  " -f $out.label, $v, $d_pct)
            if ($v -eq 0 -and $b -ne 0) { $audit = "WARN: $($out.label) zero" }
            elseif ([math]::Abs($d_pct) -gt 1.0) { $audit = "WARN: |delta|>100%" }
        }
        $ws_res.Cells.Item($row, $auditCol).Formula = $audit

        $elapsed = ((Get-Date) - $t0).TotalSeconds
        Write-Host ("  {0}({1:N1}s) {2}" -f $log_outputs, $elapsed, $audit)
    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
        WriteCell $ws_res ($idx + 13) 1 $idx
        WriteCell $ws_res ($idx + 13) 2 $s.n
        $ws_res.Cells.Item($idx + 13, $auditCol).Formula = "ERROR: $($_.Exception.Message)"
    }
}

# Restaurar baseline final
Write-Host "`nRestaurando J desde K backup..."
Invoke-COMRetry { $script:ws_in.Range("K8:K915").Copy($script:ws_in.Range("J8:J915")) | Out-Null }
$excel.CutCopyMode = $false
Invoke-COMRetry { $script:ws_in.Range("J7").Value2 = $script:CFG.BaseName }
Invoke-COMRetry { $script:ws_in.Range("F6").Value2 = $script:CFG.BaseName }
Invoke-COMRetry { $script:ws_in.Range("J$($script:CFG.IbrScenRow)").Value2 = 4 }
Invoke-COMRetry { $script:ws_in.Range("K5:K915").ClearContents() | Out-Null }
Invoke-COMRetry { $script:ws_s.Range($script:CFG.IbrSensRange).ClearContents() | Out-Null }
Invoke-COMRetry { $script:excel.CalculateFull() }
try { Invoke-COMRetry { $script:excel.Run($script:CFG.MacroName) } } catch { Write-Host "Macro final restore: $($_.Exception.Message)" }

# GUARDAR PRIMERO (datos preservados aunque format falle)
Write-Host "Guardando datos..."
$wb.Save()

# Formatos finales dinamicos (con try/catch — si falla, datos ya estan guardados)
$lastRow = 13 + $SCENARIOS.Count
try {
    foreach ($out in $OUTPUTS) {
        $col = $outColMap[$out.label]
        $valColLetter = ColLetter $col.valCol
        $pctColLetter = ColLetter $col.pctCol
        $ws_res.Range("$($valColLetter)13:$($valColLetter)$lastRow").NumberFormat = "#,##0"
        $ws_res.Range("$($pctColLetter)13:$($pctColLetter)$lastRow").NumberFormat = "0.00%"
    }
    $ws_res.Range("A12:Z$lastRow").Columns.AutoFit() | Out-Null
    $wb.Save()
} catch {
    Write-Host "Format step skipped: $($_.Exception.Message) (datos preservados)"
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
Write-Host "DONE - resultados en: $($CFG.PathDst), hoja '$($CFG.SheetResult)' filas $($CFG.ResHeaderRow)+"
