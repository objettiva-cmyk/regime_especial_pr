@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

title Motor de Arbitramento RIR70 - Sistema Operacional
set "SCRIPT=motor_arbitramento.py"
set "CONFIG=config_arbitramento_rir70.json"
set "REQ=requirements.txt"
set "RIR70_EXECUTADO_PELO_BAT=1"

echo ============================================================
echo  SISTEMA RIR70 - Motor de Arbitramento de Custo
echo ============================================================
echo  Entrada unica do operador: informe o periodo, confirme o
echo  ambiente e consulte a CAPA ao final do processamento.
echo ============================================================
echo.

if not exist "%SCRIPT%" (
  echo ERRO: arquivo %SCRIPT% nao localizado na pasta do pacote.
  echo Pasta atual: %CD%
  pause
  exit /b 1
)
if not exist "%CONFIG%" (
  echo ERRO: arquivo %CONFIG% nao localizado na pasta do pacote.
  pause
  exit /b 1
)
if not exist "%REQ%" (
  echo ERRO: arquivo %REQ% nao localizado na pasta do pacote.
  pause
  exit /b 1
)

for %%D in (output input input\movimento_item input\documentos input\inventario input\auxiliares input\xml input\ajustes input\evidencias) do (
  if not exist "%%D" mkdir "%%D"
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
  echo Instale o Python 3.10+ ou ajuste o PATH antes de executar.
  pause
  exit /b 1
)

echo [1/4] Validando dependencias...
%PYTHON_CMD% -c "import pandas, openpyxl, xlsxwriter, python_calamine" >nul 2>&1
if errorlevel 1 (
  echo Dependencias ausentes. Instalando a partir de requirements.txt...
  %PYTHON_CMD% -m pip install -r "%REQ%"
  if errorlevel 1 (
    echo ERRO: falha ao instalar dependencias. Verifique internet/proxy ou instale requirements.txt manualmente.
    pause
    exit /b 1
  )
)

echo.
echo [2/4] Periodo de processamento
echo Informe DD/MM/AAAA ou MM/AAAA. Pressione ENTER para usar o padrao.
set "DATA_INICIAL_PADRAO=01/01/2025"
set "DATA_FINAL_PADRAO=31/12/2025"
set /p "RIR70_DATA_INICIAL=Data inicial [%DATA_INICIAL_PADRAO%]: "
if not defined RIR70_DATA_INICIAL set "RIR70_DATA_INICIAL=%DATA_INICIAL_PADRAO%"
set /p "RIR70_DATA_FINAL=Data final   [%DATA_FINAL_PADRAO%]: "
if not defined RIR70_DATA_FINAL set "RIR70_DATA_FINAL=%DATA_FINAL_PADRAO%"

echo.
echo Periodo selecionado: %RIR70_DATA_INICIAL% ate %RIR70_DATA_FINAL%
echo.
echo [3/4] Executando motor...
echo.
%PYTHON_CMD% "%SCRIPT%"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo ERRO: processamento encerrado com falha. Verifique a pasta output e o log gerado.
  pause
  exit /b %RC%
)

echo.
echo [4/4] Localizando Excel final...
set "LASTFILE="
if exist "output\ultimo_arquivo_gerado.txt" (
  set /p LASTFILE=<"output\ultimo_arquivo_gerado.txt"
)
if not defined LASTFILE (
  for /f "delims=" %%F in ('dir /b /o-d "output\*.xlsx" 2^>nul') do (
    set "LASTFILE=%~dp0output\%%F"
    goto abrir_excel
  )
)

:abrir_excel
if defined LASTFILE (
  echo Abrindo Excel final: !LASTFILE!
  start "" "!LASTFILE!"
) else (
  echo AVISO: nenhum Excel final localizado em output.
)

echo.
echo Processo concluido. Consulte a CAPA: ela indica status, pendencias, links e proximo passo.
exit /b 0