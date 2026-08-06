@echo off
setlocal
cd /d "%~dp0"
call "%~dp0venv\Scripts\activate.bat"
python "%~dp0transcribir_audio.py"
pause
