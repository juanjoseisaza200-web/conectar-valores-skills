# ============================================================
#  Kit Notas de Voz - Instalador (PowerShell)
#  Replica el setup de transcripcion (Whisper small + FFmpeg).
# ============================================================
$ErrorActionPreference = "Stop"
$KitDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$DestDir    = "C:\Notas_De_Voz"
$FfmpegDir  = "C:\ffmpeg_local"
$FfmpegExe  = Join-Path $FfmpegDir "ffmpeg.exe"
$VenvDir    = Join-Path $DestDir "venv"
$PyExe      = Join-Path $VenvDir "Scripts\python.exe"
$PipExe     = Join-Path $VenvDir "Scripts\pip.exe"

function Log($msg) { Write-Host ">> $msg" -ForegroundColor Cyan }
function Ok($msg)  { Write-Host "OK $msg" -ForegroundColor Green }
function Warn($m)  { Write-Host "!! $m" -ForegroundColor Yellow }

# ---- 1. Verificar winget ---------------------------------------------------
Log "Verificando winget..."
$wg = Get-Command winget -ErrorAction SilentlyContinue
if (-not $wg) {
  Warn "winget no esta disponible. Abre Microsoft Store -> 'App Installer' y actualizalo, luego vuelve a correr este instalador."
  exit 1
}
Ok "winget presente."

# ---- 2. Instalar Python 3.14 ------------------------------------------------
Log "Buscando Python 3.14..."
$pyLauncher = Get-Command py -ErrorAction SilentlyContinue
$has314 = $false
if ($pyLauncher) {
  try {
    $vers = & py -0p 2>$null
    if ($vers -match "3\.14") { $has314 = $true }
  } catch {}
}
if (-not $has314) {
  Log "Instalando Python 3.14 via winget (acepta el UAC si aparece)..."
  winget install --id Python.Python.3.14 -e --accept-source-agreements --accept-package-agreements --silent
  if ($LASTEXITCODE -ne 0) { throw "Fallo instalacion de Python 3.14 (winget exit $LASTEXITCODE)" }
  Ok "Python 3.14 instalado."
} else {
  Ok "Python 3.14 ya estaba instalado."
}

# Resolver ruta python 3.14
$pythonHost = $null
try {
  $pythonHost = (& py -3.14 -c "import sys;print(sys.executable)") 2>$null
} catch {}
if (-not $pythonHost -or -not (Test-Path $pythonHost)) {
  # Fallback: ubicaciones tipicas
  $candidatos = @(
    "$env:LOCALAPPDATA\Programs\Python\Python314\python.exe",
    "C:\Program Files\Python314\python.exe"
  )
  foreach ($c in $candidatos) { if (Test-Path $c) { $pythonHost = $c; break } }
}
if (-not $pythonHost) { throw "No encuentro python 3.14 instalado." }
Ok "Python host: $pythonHost"

# ---- 3. Instalar FFmpeg en C:\ffmpeg_local ---------------------------------
Log "Verificando FFmpeg en $FfmpegExe ..."
if (-not (Test-Path $FfmpegExe)) {
  Log "Instalando FFmpeg via winget..."
  winget install --id Gyan.FFmpeg -e --accept-source-agreements --accept-package-agreements --silent
  if ($LASTEXITCODE -ne 0) { throw "Fallo instalacion de FFmpeg (winget exit $LASTEXITCODE)" }

  Log "Buscando ffmpeg.exe instalado..."
  $found = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Filter ffmpeg.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $found) {
    $found = Get-ChildItem -Path "C:\Program Files" -Filter ffmpeg.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
  }
  if (-not $found) { throw "No encuentro ffmpeg.exe luego de instalar." }

  if (-not (Test-Path $FfmpegDir)) { New-Item -ItemType Directory -Path $FfmpegDir | Out-Null }
  Copy-Item $found.FullName -Destination $FfmpegExe -Force
  # Copiar tambien ffprobe si existe junto
  $ffprobe = Join-Path $found.Directory "ffprobe.exe"
  if (Test-Path $ffprobe) { Copy-Item $ffprobe -Destination (Join-Path $FfmpegDir "ffprobe.exe") -Force }
  Ok "FFmpeg copiado a $FfmpegExe"
} else {
  Ok "FFmpeg ya existe en $FfmpegExe"
}

# ---- 4. Crear carpeta destino y copiar script ------------------------------
Log "Preparando $DestDir ..."
if (-not (Test-Path $DestDir)) { New-Item -ItemType Directory -Path $DestDir | Out-Null }
Copy-Item (Join-Path $KitDir "transcribir_audio.py") -Destination (Join-Path $DestDir "transcribir_audio.py") -Force
Copy-Item (Join-Path $KitDir "Notas_de_Voz.bat")     -Destination (Join-Path $DestDir "Notas_de_Voz.bat")     -Force
Copy-Item (Join-Path $KitDir "requirements.txt")     -Destination (Join-Path $DestDir "requirements.txt")     -Force
Ok "Archivos copiados a $DestDir"

# ---- 5. Crear venv ---------------------------------------------------------
if (-not (Test-Path $PyExe)) {
  Log "Creando entorno virtual (venv)..."
  & $pythonHost -m venv $VenvDir
  if ($LASTEXITCODE -ne 0) { throw "Fallo creando venv" }
  Ok "venv creado."
} else {
  Ok "venv ya existia."
}

# ---- 6. Instalar dependencias ---------------------------------------------
Log "Actualizando pip..."
& $PyExe -m pip install --upgrade pip
Log "Instalando torch (CPU)... esto puede tardar varios minutos."
& $PipExe install --index-url https://download.pytorch.org/whl/cpu torch==2.9.1
if ($LASTEXITCODE -ne 0) {
  Warn "Fallo torch desde indice CPU; intentando indice por defecto..."
  & $PipExe install torch==2.9.1
  if ($LASTEXITCODE -ne 0) { throw "No se pudo instalar torch" }
}
Log "Instalando openai-whisper y dependencias..."
& $PipExe install -r (Join-Path $DestDir "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "Fallo instalacion de requirements" }
Ok "Dependencias instaladas."

# ---- 7. Pre-descargar modelo small ----------------------------------------
Log "Pre-descargando modelo Whisper 'small' (~460 MB, una sola vez)..."
& $PyExe -c "import os; os.environ['PATH']=r'C:\ffmpeg_local'+os.pathsep+os.environ['PATH']; import whisper; whisper.load_model('small'); print('modelo OK')"
if ($LASTEXITCODE -ne 0) { Warn "Fallo pre-descarga del modelo. Se descargara en el primer uso." } else { Ok "Modelo small cacheado." }

# ---- 8. Crear acceso directo en Desktop -----------------------------------
Log "Creando acceso directo en el Escritorio..."
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "Notas de Voz.lnk"
$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut($lnkPath)
$sc.TargetPath = (Join-Path $DestDir "Notas_de_Voz.bat")
$sc.WorkingDirectory = $DestDir
$sc.IconLocation = "$env:SystemRoot\System32\SoundRecorder.exe,0"
$sc.Description = "Transcribir nota de voz a texto (Whisper)"
$sc.Save()
Ok "Acceso directo: $lnkPath"

Write-Host ""
Write-Host "Listo. Doble clic al icono 'Notas de Voz' del Escritorio." -ForegroundColor Green
exit 0
