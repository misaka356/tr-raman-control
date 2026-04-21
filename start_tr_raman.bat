@echo off
setlocal
cd /d "%~dp0"

set "PY_EXE=C:\Users\adimn\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"
set "PYW_EXE=%PY_EXE:python.exe=pythonw.exe%"
if /I "%PYW_EXE%"=="%PY_EXE%" set "PYW_EXE=pythonw"

if not exist "app_config.json" (
  copy /Y "app_config.template.json" "app_config.json" >nul
  exit /b 0
)

"%PY_EXE%" -c "import pyvisa, pyvisa_py, lmfit, numpy" >nul 2>nul
if errorlevel 1 (
  "%PY_EXE%" -m pip install pyvisa pyvisa-py psutil zeroconf lmfit
  if errorlevel 1 exit /b 1
)

start "" "%PYW_EXE%" ".\tr_raman_ui.py"
exit /b 0
