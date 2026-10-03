# Telepítés Linuxra (Arch és Debian)

Linuxon a program forrásból fut. Kell hozzá a **Python 3**, a **tkinter**
(Tk) és az **openpyxl**. Mindhármat a rendszer csomagkezelője telepíti.

## 1. A szükséges csomagok

### Arch Linux (és Arch-alapúak: CachyOS, EndeavourOS, Manjaro)

```bash
sudo pacman -S --needed python tk python-openpyxl git
```

### Debian (és Debian-alapúak: Ubuntu, Linux Mint, Raspberry Pi OS)

```bash
sudo apt update
sudo apt install python3 python3-tk python3-openpyxl git
```

> Az újabb rendszereken a `pip install` tiltva van a rendszer Pythonjában
> („externally-managed-environment”). Ezért itt az openpyxl is a
> csomagkezelővel települ.

## 2. A program letöltése

Két lehetőséged van.

**a) git-tel:**

```bash
git clone https://github.com/Botirocky/oxigen-keszletezo.git ~/oxigen-keszletezo
cd ~/oxigen-keszletezo
```

**b) kiadásból, git nélkül:** a
**https://github.com/Botirocky/oxigen-keszletezo/releases/latest** oldalon
töltsd le az `Oxigen_keszletezo_vX.Y.Z.zip` fájlt (**nem** a `windows`
nevűt), és bontsd ki:

```bash
cd ~
unzip ~/Letöltések/Oxigen_keszletezo_v*.zip    # angol rendszeren: ~/Downloads
cd ~/Oxigen_keszletezo
```

> A beépített frissítő a kiadásokból dolgozik, és a program mappájába ír. Ezért
> a b) változat a kényelmesebb. Ha git-tel töltötted le, a frissítés
> `git pull`-lal is megy.

## 3. A készlet-Excel

Másold a bolt `keszlet.xlsx` fájlját **a program mappájába** (a
`keszlet_app.py` mellé). Ha nem teszed oda, az első indításkor a program
megkérdezi, hol van, és megjegyzi.

## 4. Indítás

```bash
./inditas.sh
```

Az indító ellenőrzi, hogy megvan-e a tkinter és az openpyxl. Ha valamelyik
hiányzik, kiírja a telepítéséhez szükséges parancsot az adott disztróra.

Az **Admin** belépés alap jelszava **`admin`**. Az első belépés után
változtasd meg az *Admin → Dolgozók* fülön.

## 5. Ikon a menübe (nem kötelező)

Ha a program a `~/oxigen-keszletezo` mappában van (különben írd át az
útvonalakat):

```bash
cat > ~/.local/share/applications/oxigen-keszletezo.desktop <<EOF
[Desktop Entry]
Type=Application
Name=Oxigén készletező
Exec=$HOME/oxigen-keszletezo/inditas.sh
Path=$HOME/oxigen-keszletezo
Icon=$HOME/oxigen-keszletezo/assets/app_icon.png
Terminal=false
Categories=Office;
EOF
```

Ezután a program az alkalmazásmenüben is megjelenik.

## 6. Frissítés

- **Admin → ⬆ Frissítés:** letölti és telepíti az új verziót. A
  `keszlet.xlsx`-hez, a `config.json`-hoz és a `backup/` mappához nem nyúl.
- Vagy git-tel: `git pull` a program mappájában.

## Gyakori gondok

| Tünet | Megoldás |
|-------|----------|
| `HIBA: a tkinter (Tcl/Tk) nem érhető el` | Arch: `sudo pacman -S tk` · Debian: `sudo apt install python3-tk` |
| `ModuleNotFoundError: openpyxl` vagy `externally-managed-environment` | Arch: `sudo pacman -S python-openpyxl` · Debian: `sudo apt install python3-openpyxl` |
| `Permission denied` az `./inditas.sh`-nál | `chmod +x inditas.sh` |
| Csúnya, szögletes betűk | Telepíts egy rendes betűtípust. Arch: `sudo pacman -S noto-fonts` · Debian: `sudo apt install fonts-noto` |
