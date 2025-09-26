	@echo off
setlocal

REM 
set BASE_DIR=%~dp0
cd /d %BASE_DIR%

REM 
set VENV_DIR=.venv

REM 
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python no está instalado. Por favor instala Python 3.9+ y vuelve a intentarlo.
    pause
    exit /b 1
)


REM 
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Creando entorno virtual...
    python -m venv %VENV_DIR%
)


REM 
call "%VENV_DIR%\Scripts\activate.bat"

REM
if not exist "%VENV_DIR%\installed.txt" (
    echo Instalando dependencias...
    pip install --upgrade pip
    pip install -r requirements.txt
    echo done > "%VENV_DIR%\installed.txt"
)

REM 
echo Ejecutando aplicacion...
python -m src.GUI.gui

rem (sin 'pause')