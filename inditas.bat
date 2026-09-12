@echo off
chcp 65001 >nul
REM ============================================================
REM  Keszletkezelo - Eladas (Excel-alapu kassza)
REM  Dupla kattintassal vagy parancssorbol futtathato.
REM ============================================================
setlocal

REM A szkript sajat mappajaba lepunk (igy barhonnan futtathato)
cd /d "%~dp0"

REM Python megkeresese: eloszor a "py" launcher, aztan a "python"
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)

if not defined PY (
    echo [HIBA] Nem talalhato a Python.
    echo Telepitsd innen: https://www.python.org/downloads/
    echo Telepiteskor pipald be a "Add Python to PATH" opciot!
    echo.
    pause
    exit /b 1
)

REM openpyxl megletenek ellenorzese; ha hianyzik, telepitjuk
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Az openpyxl nincs telepitve - telepites folyamatban...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [HIBA] A fuggosegek telepitese nem sikerult.
        pause
        exit /b 1
    )
)

REM Alkalmazas inditasa
%PY% keszlet_app.py
if errorlevel 1 (
    echo.
    echo [HIBA] Az alkalmazas hibaval lepett ki.
    pause
    exit /b 1
)

endlocal
