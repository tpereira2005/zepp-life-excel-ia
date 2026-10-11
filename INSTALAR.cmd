@echo off
rem Cria o ambiente Python do projeto (.venv) e instala as dependencias fixadas.
setlocal
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv || goto erro
)
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt || goto erro
if not exist "config.json" copy /y "config.example.json" "config.json" >nul && echo Criado config.json: ajuste a data de inicio e o protocolo.
if not exist "bin\zepp-reader.exe" echo Aviso: falta bin\zepp-reader.exe. Veja docs\INSTALACAO.md.
echo Instalacao concluida.
pause
exit /b 0
:erro
echo A instalacao falhou. Confirme que o Python 3.11 ou superior esta instalado.
pause
exit /b 1
