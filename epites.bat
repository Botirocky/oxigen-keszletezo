@echo off
chcp 65001 >nul
REM ============================================================
REM  Oxigen keszletezo - onallo EXE keszitese (Windows)
REM  Ez a fajl LETOLT MINDEN szukseges elemet egy sajat virtualis
REM  kornyezetbe (.venv-build), igy a rendszer Pythonjat nem
REM  piszkalja, majd egyetlen onallo .exe-t epit a dist mappaba.
REM  Dupla kattintassal futtathato - eleg egyszer lefuttatni.
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

REM 1) Virtualis kornyezet letrehozasa (ha meg nincs)
set "VENV=.venv-build"
set "VPY=%VENV%\Scripts\python.exe"
if not exist "%VPY%" (
    echo Virtualis kornyezet letrehozasa (%VENV%)...
    %PY% -m venv "%VENV%"
    if errorlevel 1 (
        echo [HIBA] A venv letrehozasa nem sikerult.
        pause
        exit /b 1
    )
)

REM 2) MINDEN szukseges elem letoltese a venv-be: pip frissites + fuggosegek + PyInstaller
echo.
echo A szukseges csomagok letoltese es telepitese a venv-be...
echo    - pip (frissites)
echo    - openpyxl (Excel-kezeles)
echo    - pyinstaller (a build eszkoze)
"%VPY%" -m pip install --upgrade pip
if errorlevel 1 (
    echo [HIBA] A pip frissitese nem sikerult (ellenorizd az internetkapcsolatot).
    pause
    exit /b 1
)
"%VPY%" -m pip install --upgrade -r requirements.txt pyinstaller
if errorlevel 1 (
    echo [HIBA] A csomagok letoltese/telepitese nem sikerult.
    pause
    exit /b 1
)

REM 3) Korabbi build maradvanyok takaritasa
echo Korabbi build takaritasa...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Oxigen_keszletezo.spec" del /q "Oxigen_keszletezo.spec"

REM 4) EXE epitese: egyetlen fajl, ablakos mod, assets becsomagolva, sajat ikon
echo.
echo EXE epitese folyamatban (ez eltarthat egy percig)...
echo.
"%VPY%" -m PyInstaller --onefile --windowed --name "Oxigen_keszletezo" ^
  --add-data "assets;assets" ^
  --icon "assets\app_icon.ico" ^
  keszlet_app.py
if errorlevel 1 (
    echo.
    echo [HIBA] A build nem sikerult.
    pause
    exit /b 1
)

REM 5) A keszlet.xlsx odamasolasa az exe melle, hogy a csomag azonnal hasznalhato legyen
if exist "keszlet.xlsx" (
    copy /y "keszlet.xlsx" "dist\keszlet.xlsx" >nul
    echo A keszlet.xlsx az exe melle masolva.
)

echo.
echo ============================================================
echo  KESZ!  Az elkeszult program:  dist\Oxigen_keszletezo.exe
echo.
echo  Telepiteshez masold a dist mappa TARTALMAT a bolti gepre.
echo  A keszlet.xlsx-et az exe mellett szerkesztheti a bolt.
echo ============================================================
echo.
pause

endlocal
