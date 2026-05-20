@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

title Sistema Web RIR70
set "SCRIPT=sistema_rir70_web.py"

if not exist "%SCRIPT%" (
  echo ERRO: arquivo %SCRIPT% nao localizado.
  pause
  exit /b 1
)

set "PYTHON_CMD="
python --version >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=python"
if not defined PYTHON_CMD (
  py -3 --version >nul 2>&1
  if not errorlevel 1 set "PYTHON_CMD=py -3"
)
if not defined PYTHON_CMD (
  echo ERRO: Python 3 nao localizado no PATH.
  pause
  exit /b 1
)

echo Abrindo Sistema RIR70 no navegador...
echo Mantenha esta janela aberta enquanto usa o sistema.
echo.
%PYTHON_CMD% "%SCRIPT%"
