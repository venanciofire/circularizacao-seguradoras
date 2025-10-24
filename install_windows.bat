@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

REM ==============================================
REM  Instalador do Pipeline de Circularização Apólices (Windows)
REM  - Cria venv .venv
REM  - Instala dependências
REM  - Gera atalho de Desktop para a GUI
REM  - Mesa de Operações e Bko
REM  - Autor: Venâncio - Carlos Venâncio
REM ==============================================

REM Ir para a pasta do script
pushd "%~dp0"
set "ROOT=%~dp0"

REM Descobrir Python
set "PYTHON="
where py >nul 2>&1 && set "PYTHON=py -3"
if not defined PYTHON (
  where python >nul 2>&1 && set "PYTHON=python"
)
if not defined PYTHON (
  echo [ERRO] Python 3.9+ nao encontrado no PATH (py ou python).
  echo Instale o Python (https://www.python.org/downloads/) e tente novamente.
  pause
  exit /b 1
)

REM Verificar versao >= 3.9
%PYTHON% -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3,9) else 1)"
if errorlevel 1 (
  echo [ERRO] E necessario Python 3.9 ou superior.
  pause
  exit /b 1
)

echo.
echo [1/4] Criando ambiente virtual (.venv)...
%PYTHON% -m venv .venv
if errorlevel 1 (
  echo [ERRO] Falha ao criar o ambiente virtual.
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERRO] Falha ao ativar o ambiente virtual.
  pause
  exit /b 1
)

echo.
echo [2/4] Atualizando pip...
python -m pip install --upgrade pip
if errorlevel 1 (
  echo [AVISO] Nao foi possivel atualizar o pip. Prosseguindo...
)

echo.
echo [3/4] Instalando dependencias...
if exist requirements.txt (
  python -m pip install -r requirements.txt
) else (
  python -m pip install pandas openpyxl xlrd
)
if errorlevel 1 (
  echo [ERRO] Falha na instalacao de dependencias.
  pause
  exit /b 1
)

echo.
echo [4/4] Validando instalacao...
python - <<PYEND
try:
    import pandas, openpyxl, xlrd
    print('[OK] Dependencias carregadas com sucesso.')
except Exception as e:
    print('[ERRO] Falha ao importar dependencias:', e)
    raise
PYEND
if errorlevel 1 (
  echo [ERRO] Validacao falhou.
  pause
  exit /b 1
)

REM Criar run_gui.bat se ainda nao existir
if not exist run_gui.bat (
  > run_gui.bat (
    echo @echo off
    echo setlocal
    echo pushd "%%~dp0"
    echo call ".venv\Scripts\activate.bat"
    echo python gui_processa_seguradoras.py
    echo endlocal
  )
)

REM Criar atalho no Desktop via PowerShell (opcional)
set "DESK=%UserProfile%\Desktop"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$root=$env:ROOT; $desk=$env:DESK; $lnk=Join-Path $desk 'Circularização Apolices (GUI).lnk'; ^
   $w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut($lnk); ^
   $s.TargetPath = (Join-Path $root 'run_gui.bat'); ^
   $s.WorkingDirectory = $root; ^
   $s.IconLocation = [System.Environment]::SystemDirectory + '\\shell32.dll,2'; ^
   $s.Save()" >nul 2>&1

if exist "%DESK%\Circularização Apolices (GUI).lnk" (
  echo.
  echo [SUCESSO] Instalacao concluida. Atalho criado no Desktop: Pipeline Apolices (GUI)
) else (
  echo.
  echo [SUCESSO] Instalacao concluida. Se preferir, execute run_gui.bat para abrir a interface.
)

echo.
echo Para executar agora a GUI, pressione qualquer tecla...
pause >nul
start "" run_gui.bat

endlocal
