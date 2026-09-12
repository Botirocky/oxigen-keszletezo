#!/usr/bin/env bash
# Oxigén készletező – önálló futtatható fájl építése (Linux/macOS, PyInstaller)
# Ez a szkript LETÖLT MINDEN szükséges elemet egy saját virtuális környezetbe
# (.venv-build), így a rendszer Pythonját nem piszkálja, majd egyetlen önálló
# binárist épít. Modern disztrók (Arch, Debian 12+, Fedora) "externally managed"
# pip-jét is kezeli, mert mindent a venv-be telepít.
set -u
cd "$(dirname "$0")" || exit 1

PY="${PYTHON:-python3}"
VENV=".venv-build"

# 1) Python ellenőrzése
if ! command -v "$PY" >/dev/null 2>&1; then
    echo "HIBA: nincs telepítve a python3." >&2
    exit 1
fi

# 2) tkinter (Tcl/Tk) ellenőrzése – RENDSZERCSOMAG, pip-pel NEM telepíthető
if ! "$PY" -c "import tkinter" >/dev/null 2>&1; then
    echo "HIBA: a tkinter (Tcl/Tk) nem érhető el – a build is ezt igényli." >&2
    echo "Telepítsd a rendszer csomagkezelőjével:" >&2
    if command -v pacman >/dev/null 2>&1; then
        echo "    sudo pacman -S tk" >&2
    elif command -v apt >/dev/null 2>&1; then
        echo "    sudo apt install python3-tk python3-venv" >&2
    elif command -v dnf >/dev/null 2>&1; then
        echo "    sudo dnf install python3-tkinter" >&2
    else
        echo "    (telepítsd a 'tk' vagy 'python3-tk' csomagot)" >&2
    fi
    exit 1
fi

# 3) Virtuális környezet létrehozása (ha még nincs)
if [ ! -x "$VENV/bin/python" ]; then
    echo "Virtuális környezet létrehozása ($VENV)..."
    "$PY" -m venv "$VENV" || {
        echo "HIBA: a venv létrehozása nem sikerült." >&2
        echo "Debian/Ubuntu: sudo apt install python3-venv" >&2
        exit 1
    }
fi
VPY="$VENV/bin/python"

# 4) MINDEN szükséges elem letöltése a venv-be: pip frissítés + függőségek + PyInstaller
echo
echo "A szükséges csomagok letöltése és telepítése a venv-be..."
echo "   - pip (frissítés)"
echo "   - openpyxl (Excel-kezelés)"
echo "   - pyinstaller (a build eszköze)"
"$VPY" -m pip install --upgrade pip || { echo "HIBA: a pip frissítése nem sikerült." >&2; exit 1; }
"$VPY" -m pip install --upgrade -r requirements.txt pyinstaller \
    || { echo "HIBA: a csomagok letöltése/telepítése nem sikerült." >&2; exit 1; }

# 5) korábbi build maradványok takarítása
echo "Korábbi build takarítása..."
rm -rf build dist Oxigen_keszletezo.spec

# 6) EXE/bináris építése: egyetlen fájl, ablakos mód, assets becsomagolva
#    (a Linux figyelmen kívül hagyja az ikont, ezért nem adunk --icon-t)
echo
echo "Build folyamatban (ez eltarthat egy percig)..."
echo
"$VPY" -m PyInstaller --onefile --windowed --name "Oxigen_keszletezo" \
    --add-data "assets:assets" \
    keszlet_app.py || { echo "HIBA: a build nem sikerült." >&2; exit 1; }

# 7) a keszlet.xlsx a bináris mellé, hogy a csomag azonnal használható legyen
if [ -f keszlet.xlsx ]; then
    cp -f keszlet.xlsx dist/keszlet.xlsx
    echo "A keszlet.xlsx a bináris mellé másolva."
fi

echo
echo "============================================================"
echo " KÉSZ!  Az elkészült program:  dist/Oxigen_keszletezo"
echo
echo " Indítás:        ./dist/Oxigen_keszletezo"
echo " A keszlet.xlsx-et a bináris mellett szerkesztheted."
echo "============================================================"
