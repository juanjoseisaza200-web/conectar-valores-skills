@echo off
REM ============================================================
REM  Kit Notas de Voz - Instalador automatico
REM  Replica el sistema de transcripcion de David en este PC.
REM ============================================================
cd /d "%~dp0"
echo.
echo ============================================================
echo   INSTALADOR - Sistema de transcripcion de notas de voz
echo ============================================================
echo.
echo Este proceso instalara: Python 3.14, FFmpeg y Whisper.
echo Puede tardar 10-20 minutos la primera vez (descarga ~2 GB).
echo.
pause

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1"
set ERR=%ERRORLEVEL%

echo.
if %ERR% EQU 0 (
  echo ============================================================
  echo   INSTALACION COMPLETADA
  echo ============================================================
  echo.
  echo Tienes un acceso directo "Notas de Voz" en tu Escritorio.
  echo Doble clic ahi -^> selecciona el audio -^> espera la transcripcion.
) else (
  echo ============================================================
  echo   ERROR EN LA INSTALACION (codigo %ERR%)
  echo ============================================================
  echo Revisa los mensajes arriba o avisa a David con la captura.
)
echo.
pause
