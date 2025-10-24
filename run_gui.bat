@echo off
setlocal
pushd "%~dp0"
if exist .venv\Scripts\activate.bat (
  call ".venv\Scripts\activate.bat"
)
python gui_processa_seguradoras.py
endlocal
