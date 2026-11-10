@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  where py >nul 2>nul
  if errorlevel 1 (
    python -m venv .venv
  ) else (
    py -3 -m venv .venv
  )
  if errorlevel 1 goto error
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto error
".venv\Scripts\python.exe" preparar_carga.py %*
if errorlevel 1 goto error
echo.
echo Listo. Revisa la carpeta output.
pause
exit /b 0
:error
echo.
echo No se completo el proceso. Revisa el mensaje anterior.
echo Se necesita Python 3.10 o superior e Internet para instalar las dependencias.
pause
exit /b 1
