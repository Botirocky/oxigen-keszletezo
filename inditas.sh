#!/usr/bin/env bash
# Oxigén – Eladás indítása (Linux/macOS)
set -u
cd "$(dirname "$0")" || exit 1

PY="${PYTHON:-python3}"

# 1) Python ellenőrzése
if ! command -v "$PY" >/dev/null 2>&1; then
    echo "HIBA: nincs telepítve a python3." >&2
    exit 1
fi

# 2) tkinter (Tcl/Tk) ellenőrzése – ez RENDSZERCSOMAG, pip-pel NEM telepíthető
if ! "$PY" -c "import tkinter" >/dev/null 2>&1; then
    echo "HIBA: a tkinter (Tcl/Tk) nem érhető el." >&2
    echo "Telepítsd a rendszer csomagkezelőjével:" >&2
    if command -v pacman >/dev/null 2>&1; then
        echo "    sudo pacman -S tk" >&2
    elif command -v apt >/dev/null 2>&1; then
        echo "    sudo apt install python3-tk" >&2
    elif command -v dnf >/dev/null 2>&1; then
        echo "    sudo dnf install python3-tkinter" >&2
    else
        echo "    (telepítsd a 'tk' vagy 'python3-tk' csomagot)" >&2
    fi
    exit 1
fi

# 3) openpyxl telepítése, ha hiányzik
if ! "$PY" -c "import openpyxl" 2>/dev/null; then
    if ! "$PY" -m pip install --user -r requirements.txt; then
        echo "HIBA: az openpyxl nem települt. Telepítsd a rendszer csomagkezelőjével:" >&2
        if command -v pacman >/dev/null 2>&1; then
            echo "    sudo pacman -S python-openpyxl" >&2
        elif command -v apt >/dev/null 2>&1; then
            echo "    sudo apt install python3-openpyxl" >&2
        elif command -v dnf >/dev/null 2>&1; then
            echo "    sudo dnf install python3-openpyxl" >&2
        fi
        exit 1
    fi
fi

# 4) indítás
exec "$PY" keszlet_app.py
