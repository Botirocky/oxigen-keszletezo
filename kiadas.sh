#!/usr/bin/env bash
#
# Kiadás készítése – ezzel tudsz új verziót kiadni a boltnak
# ==========================================================
# Mit csinál:
#   1. felemeli a verziószámot a frissito.py-ban
#   2. összecsomagolja a programot egy zip-be
#   3. létrehozza a GitHub-kiadást (release), ha van 'gh' parancs
#
# A bolti gépen futó program ezt a kiadást fogja meglátni és feltelepíteni.
#
# Használat:
#   ./kiadas.sh 1.2.0            – pontos verziószám
#   ./kiadas.sh --javitas        – 1.1.0 -> 1.1.1  (hibajavítás)
#   ./kiadas.sh --funkcio        – 1.1.0 -> 1.2.0  (új funkció)
#   ./kiadas.sh 1.2.0 --csak-zip – csak a zip készül el, feltöltés nélkül
#   ./kiadas.sh 1.2.0 -m "Szöveg" – saját kiadási leírás
#
set -euo pipefail
cd "$(dirname "$0")"

PIROS=$'\e[31m'; ZOLD=$'\e[32m'; SARGA=$'\e[33m'; SZURKE=$'\e[90m'; V=$'\e[0m'
hiba() { echo "${PIROS}HIBA:${V} $*" >&2; exit 1; }
ok()   { echo "${ZOLD}✔${V} $*"; }
info() { echo "  $*"; }

MOST="$(sed -n 's/^APP_VERSION = "\(.*\)"/\1/p' frissito.py | head -1)"
[[ -n "$MOST" ]] || hiba "nem találom az APP_VERSION-t a frissito.py-ban"

UJ=""; CSAK_ZIP=0; LEIRAS=""; SABLONNAL=0
while [[ $# -gt 0 ]]; do
    case "$1" in
        --javitas) UJ="$(awk -F. '{printf "%d.%d.%d", $1, $2, $3+1}' <<<"$MOST")"; shift ;;
        --funkcio) UJ="$(awk -F. '{printf "%d.%d.0", $1, $2+1}' <<<"$MOST")"; shift ;;
        --nagy)    UJ="$(awk -F. '{printf "%d.0.0", $1+1}' <<<"$MOST")"; shift ;;
        --csak-zip) CSAK_ZIP=1; shift ;;
        --sablonnal) SABLONNAL=1; shift ;;
        -m|--uzenet) LEIRAS="${2:-}"; shift 2 ;;
        -h|--help) sed -n '3,20p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
        *) UJ="$1"; shift ;;
    esac
done

[[ -n "$UJ" ]] || hiba "add meg az új verziót (pl. ./kiadas.sh 1.2.0) vagy használd a --javitas / --funkcio kapcsolót"
[[ "$UJ" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || hiba "a verzió alakja X.Y.Z legyen (most: $UJ)"

echo
echo "  Jelenlegi verzió : ${MOST}"
echo "  Új verzió        : ${ZOLD}${UJ}${V}"
echo

# --- 1) ellenőrzés: fordul-e a program? --------------------------------------
python3 -c "import ast,sys; ast.parse(open('keszlet_app.py',encoding='utf-8').read())" \
    || hiba "a keszlet_app.py szintaktikailag hibás – ezt nem adjuk ki"
python3 -c "import ast,sys; ast.parse(open('frissito.py',encoding='utf-8').read())" \
    || hiba "a frissito.py szintaktikailag hibás"
ok "a program fordul"

# --- 2) verziószám emelése ---------------------------------------------------
sed -i "s/^APP_VERSION = \".*\"/APP_VERSION = \"${UJ}\"/" frissito.py
ok "verziószám átírva a frissito.py-ban"

# --- 3) csomagolás -----------------------------------------------------------
CSOMAG="Oxigen_keszletezo"
ZIP="${CSOMAG}_v${UJ}.zip"
MUNKA="$(mktemp -d)"
trap 'rm -rf "$MUNKA"' EXIT
CEL="${MUNKA}/${CSOMAG}"
mkdir -p "$CEL"

for f in keszlet_app.py frissito.py requirements.txt README.md \
         inditas.bat inditas.sh epites.bat epites.sh; do
    [[ -f "$f" ]] && cp "$f" "$CEL/"
done
[[ -d assets ]] && cp -r assets "$CEL/"

if (( SABLONNAL )); then
    # ÜRES sablon – soha ne a bolt éles adatait!
    if [[ -f keszlet_teszt.xlsx ]]; then
        cp keszlet_teszt.xlsx "$CEL/keszlet.xlsx"
        ok "üres sablon-Excel becsomagolva (keszlet_teszt.xlsx)"
    else
        echo "${SARGA}!${V} nincs keszlet_teszt.xlsx – Excel nélkül csomagolok"
    fi
fi

cat >"$CEL/OLVASD_EL.txt" <<OLVASS
Oxigén készletező – v${UJ}
$(printf '=%.0s' {1..40})

Indítás Windowson:  kattints az inditas.bat fájlra.
Indítás Linuxon:    ./inditas.sh

FONTOS: a keszlet.xlsx fájlt NE töröld – az a program adatbázisa.
A frissítés soha nem írja felül sem az Excelt, sem a beállításokat.

Kiadva: $(date '+%Y-%m-%d')
OLVASS

( cd "$MUNKA" && zip -qr "$OLDPWD/$ZIP" "$CSOMAG" )
ok "csomag kész: ${ZIP} ($(du -h "$ZIP" | cut -f1))"

# a csomag ellenőrzése: tényleg benne van-e, aminek kell
python3 - "$ZIP" <<'ELLENOR'
import sys, zipfile
z = zipfile.ZipFile(sys.argv[1])
nevek = z.namelist()
kell = ["keszlet_app.py", "frissito.py"]
hiany = [k for k in kell if not any(n.endswith("/" + k) for n in nevek)]
if hiany:
    print(f"HIBA: hiányzik a csomagból: {hiany}"); sys.exit(1)
veszely = [n for n in nevek if n.endswith("config.json")]
if veszely:
    print(f"HIBA: a csomagban van config.json (jelszó-hash!): {veszely}"); sys.exit(1)
print(f"  a csomag {len(nevek)} elemet tartalmaz, az ellenőrzés rendben")
ELLENOR
ok "a csomag ellenőrizve"

if (( CSAK_ZIP )); then
    echo
    echo "  Csak a zip készült el. Feltöltés kézzel:"
    echo "    ${SZURKE}GitHub → Releases → Draft a new release → tag: v${UJ} → a zip feltöltése${V}"
    exit 0
fi

# --- 4) GitHub kiadás --------------------------------------------------------
[[ -n "$LEIRAS" ]] || LEIRAS="Oxigén készletező v${UJ}"

if command -v gh >/dev/null 2>&1; then
    if [[ -d .git ]]; then
        git add -A && git commit -qm "v${UJ}" || true
        # ANNOTÁLT címke kell: a könnyű címkét a --follow-tags nem tolja fel.
        git tag -fa "v${UJ}" -m "v${UJ}" >/dev/null
        git push -q origin HEAD 2>/dev/null || echo "${SARGA}!${V} a git push nem sikerült"
        git push -qf origin "v${UJ}" 2>/dev/null || echo "${SARGA}!${V} a címke feltöltése nem sikerült"
    fi
    gh release create "v${UJ}" "$ZIP" --title "v${UJ}" --notes "$LEIRAS" \
        && ok "GitHub-kiadás létrehozva: v${UJ}" \
        || hiba "a gh release nem sikerült"
    echo
    echo "  A bolti gép a következő indításkor meglátja az új verziót."
else
    echo
    echo "${SARGA}!${V} Nincs telepítve a 'gh' parancs, ezért a feltöltés kimarad."
    echo "  Két lehetőség:"
    echo "    a) telepítsd:  sudo pacman -S github-cli  (majd: gh auth login)"
    echo "    b) töltsd fel kézzel: GitHub → Releases → Draft a new release"
    echo "       tag: ${ZOLD}v${UJ}${V}  ·  csatolt fájl: ${ZOLD}${ZIP}${V}"
fi
