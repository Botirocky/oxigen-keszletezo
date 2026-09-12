# Oxigén készletező (Excel-alapú kassza)

Egyszerű, vonalkódolvasóra optimalizált felület az **Oxigén Cipőbolt** számára,
**két üzemmóddal**:

- **🛒 Eladás** – vásárlónak történő eladás (csökkenti a készletet, a forgalomba számít).
- **📦 Elvitel** – eladáson kívüli készletcsökkenés: **sérült/selejt**, **visszáru
  beszállítónak**, **saját/egyéb**. Külön nyilvántartva, hogy a forgalmi
  statisztika tiszta maradjon, és minden elvitel **indokkal, dátummal a `Napló`
  lapra** is bekerül.

Az **adatbázis maga az Excel-fájl** (`keszlet.xlsx`), így bármikor megnyithatod
és szerkesztheted Excelben (árak, új termékek, statisztika stb.) – a program
ugyanazt a fájlt használja.

A felület a bolt logójára épülő, letisztult zöld arculatot kapott (elvitel
módban borostyán színűre vált), finom animációkkal (az összeg „felpörög”, a
rögzítő gomb lágyan lélegzik, indításkor beúszik az ablak).

A program **nem írja át** a táblázatod szerkezetét: csak az adott hónaphoz
tartozó **eladás** oszlopba írja be az eladott darabszámot, a képletek (Σ,
KÉSZLET, KÉSZLET ÉRTÉK) maguktól újraszámolódnak.

---

## Belépés – Dolgozói és Admin felület

Indításkor egy **belépő képernyő** jelenik meg, ahol kiválasztod, milyen
módban szeretnél dolgozni:

- **🛒 Dolgozó (kassza)** – a megszokott eladási/elvitel felület. Ha vannak
  felvett dolgozók (lásd Admin → Dolgozók), belépéskor kiválasztod a neved, és
  az **eladás-naplóba bekerül, ki adta el** a terméket.
- **🔑 Admin (kezelés)** – jelszóval védett kezelői felület. **Az alap jelszó:
  `admin`** – ezt az Admin → *Dolgozók* fülön bármikor megváltoztathatod.

A felület tetején a **„Kijelentkezés”** gombbal bármikor visszatérhetsz a
belépő képernyőre (másik dolgozó / admin belépéséhez), kilépés nélkül.

### Az Admin felület füllapjai

| Fül | Mit tud |
|-----|---------|
| **📦 Termékek** | Termékek listája kereséssel; **új termék felvétele** és meglévő **szerkesztése** (név, méret, vonalkód, eladási ár, nyitó készlet) – közvetlenül az Excelbe. |
| **📊 Statisztika** | Termékszám, összes készlet, **készletérték**, éves eladott db és **bevétel**; havi forgalom-bontás; legkelendőbb termékek; **alacsony készlet** lista. |
| **📅 Napi zárás** | Egy adott nap **kassza-összesítője**: napi bevétel, **nyugták száma**, eladott db, átlagos kosárérték és aznapi elvitel; **eladónkénti bontás**, aznap eladott termékek és elvitelek. A zárás **kinyomtatható/menthető** szövegfájlba. |
| **📒 Napló** | Az **Eladások** (ki, mit, mennyiért, mikor) és az **Elvitelek** naplójának böngészése, kereshetően. |
| **👥 Dolgozók** | Dolgozók (eladók) **felvétele/törlése**, és az **admin jelszó** módosítása. |
| **⬆ Frissítés** | A program **új verziójának** keresése és telepítése (lásd lent). Ha van újabb, a fül feliratán egy pötty jelzi. |

> A dolgozók névsorát és az admin jelszó (hash-ét) a `config.json` tárolja –
> érzékeny adat nem kerül olvasható formában a fájlba.
> Az eladásokat a program az Excel **`Eladások`** munkalapjára naplózza
> (a lapot az első eladásnál automatikusan létrehozza).

### Napi zárás (kassza-összesítő)

A **📅 Napi zárás** fül a nap végi elszámolást segíti. Fent kiválasztod a
**napot** (alapból a mai), és rögtön látod:

- **Bevétel** – az adott nap összes eladásának összege,
- **Nyugták száma** – hány külön eladást (tranzakciót) rögzítettél,
- **Eladott db** és **átlagos kosárérték**,
- **Elvitel** – aznapi, eladáson kívüli készletcsökkenés darabszáma.

Alatta három bontás: **eladónként** (ki mennyit adott el), az **aznap eladott
termékek**, és az **aznapi elvitelek**. A **„💾 Zárás mentése…”** gombbal a
napi zárás egy szépen tördelt, **kinyomtatható szövegfájlba** menthető (pl. a
napi kassza mellé, papíron vagy archívumba).

> A napi zárás mindig **frissen olvassa** az Excel `Eladások` és `Napló`
> lapjait, tehát valós időben tükrözi az aznapi forgalmat. A nyugtákat az
> eladások időbélyege alapján különbözteti meg (másodperc pontosságig).

---

## Telepítés

Két módon telepítheted az appot – válaszd, ami a helyzetedhez illik.

### A) Önálló `.exe` készítése (ajánlott a bolti gépre)

Így a bolti gépre **nem kell Python**, csak egy kész program.

1. A **fejlesztői gépen** (ahol van Python) dupla kattintás az **`epites.bat`**
   fájlra (Linux/macOS: `./epites.sh`).
2. A szkript **mindent letölt és beállít automatikusan**: frissíti a `pip`-et,
   telepíti a szükséges csomagokat (`openpyxl`) **és a `PyInstaller`-t**, majd
   egyetlen önálló programot épít.
3. Ha kész, az elkészült program a **`dist`** mappában lesz:
   `dist\Oxigen_keszletezo.exe` (a `keszlet.xlsx` is odakerül mellé).
4. **Telepítés a bolti gépre:** másold át a `dist` mappa **teljes tartalmát** a
   bolti gépre (pl. az Asztalra). Indítás: dupla kattintás az
   `Oxigen_keszletezo.exe`-re. A `keszlet.xlsx`-et az exe mellett szerkesztheted.

> A `epites.bat`-ot elég **egyszer** lefuttatni. Csak akkor kell újra, ha a
> program kódja változott és új `.exe`-t szeretnél.

### B) Futtatás Pythonból (forrásból)

Ha a gépen van Python, közvetlenül is futtathatod:

- **Windows:** dupla kattintás az **`inditas.bat`** fájlra. Az első indításkor
  automatikusan telepíti a szükséges `openpyxl` csomagot.
- **Linux/macOS:** `./inditas.sh`
- **Bármilyen rendszeren, parancssorból:**

```bat
python keszlet_app.py
```

(Ehhez a `pip install -r requirements.txt` paranccsal telepítsd egyszer a
függőséget – ezt az indító szkriptek maguktól megteszik.)

---

## Indítás

Az első induláskor a program megkeresi a `keszlet.xlsx` fájlt (a saját
mappájában és a Letöltések mappában). Ha nem találja, **kiválaszthatod**.
A választást megjegyzi (`config.json`), legközelebb már nem kérdez.

> Ha máshová akarod tenni az Excelt, a program alján a **„Excel kiválasztása…”**
> gombbal bármikor átállíthatod.

---

## Használat (eladás) – lépésről lépésre

1. Indítsd el a programot. A kurzor a **beolvasó mezőben** villog.
2. **Olvasd be a terméket** a vonalkódolvasóval. A program kiírja a nevét és az
   árát, pl. `✓ 73690 / 38   27 990 Ft`, és felveszi a listára.
3. **Olvass be több terméket** is egymás után. Ha ugyanazt a terméket olvasod be
   kétszer, a darabszám automatikusan nő (×2, ×3 …).
4. Amikor minden tételt beolvastál, **nyomj Entert úgy, hogy a beolvasó mező
   üres** (vagy kattints a zöld **„Eladás rögzítése”** gombra).
5. Megjelenik az **összeg** és a tételek listája. A vásárló fizet, te
   megerősíted (**Igen**).
6. A program **beírja az eladást az Excelbe** és **levonja a készletből**.
   A lista kiürül, jöhet a következő vásárló.

### Gyakorlati példa
A vásárló két cipőt vesz:

- Beolvasod az elsőt → `✓ 73690 / 38   27 990 Ft`
- Beolvasod a másodikat → `✓ 117730 / 37   22 990 Ft`
- Üres mezőben **Enter** → „Összesen (2 db): **50 980 Ft**” → **Igen** →
  rögzítve, készlet levonva.

---

## Mit csinál hiba esetén?

- **Ismeretlen vonalkód** (nincs az Excelben): piros figyelmeztetés + hangjelzés,
  a terméket nem veszi fel. (Vedd fel előbb Excelben a terméket a vonalkóddal.)
- **Az Excel meg van nyitva** rögzítéskor: a program szól, hogy zárd be az
  Excelt, és próbáld újra. **A kosár közben megmarad**, nem veszik el.
- **Elfogyott készlet**: a tételnél a „Készlet utána” pirosan mutatja, ha
  mínuszba menne, de az eladást nem tiltja (lehet, hogy fizikailag van készlet,
  csak a nyilvántartás pontatlan).

---

## Tételek javítása eladás közben

- **− 1 db**: a kijelölt sorból egy darabot levesz (vagy dupla kattintás a soron).
- **🗑 Tétel törlése**: a kijelölt sort teljesen törli (vagy `Delete` billentyű).
- **Új eladás (ürítés)**: az egész kosarat törli (megerősítéssel).

A jobb felső sarokban a **hónap** állítható – alapból az aktuális hónap. Csak
akkor állítsd át, ha egy korábbi hónaphoz akarsz eladást utólag rögzíteni.

---

## Elvitel (eladáson kívüli készletcsökkenés)

A felület tetején válts át **📦 Elvitel** módba (a felület borostyán színűre
vált, hogy ne keverd össze az eladással). A használat ugyanaz: beolvasod a
kivett termékeket, majd üres mezőben **Enter** (vagy a **„Elvitel rögzítése”**
gomb). Rögzítés előtt kiválasztod az **okot**:

- **Sérült / selejt** – tönkrement, nem eladható áru
- **Visszáru beszállítónak** – visszaküldöd a forgalmazónak
- **Saját / egyéb** – saját célra, bemutató, ajándék stb.

Az elvitel az `ÖSSZES` lap **`elvitel` oszlopába** kerül (az aktuális hónaphoz),
és külön bejegyzésként a **`Napló`** lapra is (dátum, vonalkód, név, méret, db,
indok). A `KÉSZLET` ettől automatikusan csökken.

> Az `elvitel` oszlopokat és a `Napló` lapot a program **automatikusan
> létrehozza** az első elvitelnél, ha még nincsenek – kézzel nem kell semmit
> előkészíteni.

Mód váltása a kosarat üríti (megerősítéssel, ha van benne tétel).

---

## Az Excel szerkezete

A program a **`ÖSSZES`** munkalapot használja. Az oszlopokat a fejléc-feliratok
alapján ismeri fel, tehát ha bővíted a táblát, ezek a feliratok maradjanak meg:

| Oszlop | Fejléc | Szerep |
|--------|--------|--------|
| A | Megnevezés | termék neve |
| B | Egyéb | méret / szín / íz … |
| C | **Vonalkód** | a kereséshez ez kell |
| D | Forgalmazó | beszállító |
| E | Áfa | áfakulcs (%) |
| F | Nettó beszerzés | nettó beszerzési ár |
| G | Bruttó beszerzés | képlet |
| H | **Bruttó eladás** | ez az ár jelenik meg a kasszán |
| I | Nyitó készlet | január 1-i készlet |
| J–U | beszerzés 1–12 | havi beszerzett db |
| V | Σ | beszerzés összege (képlet) |
| W–AH | **eladás 1–12** | havi eladott db – **ide ír az Eladás mód** |
| AI | Σ | eladás összege (képlet) |
| AJ | KÉSZLET | `= Nyitó + beszerzés − eladás − elvitel` (képlet) |
| AK | KÉSZLET ÉRTÉK | `= készlet × nettó beszerzés` (képlet) |
| AL–AW | **elvitel 1–12** | havi elvitt db – **ide ír az Elvitel mód** (a program hozza létre) |
| AX | Σ | elvitel összege (képlet) |

A program egy **`Napló`** munkalapot is létrehoz az elvitelek részletes
naplózásához (dátum, vonalkód, megnevezés, méret, darab, indok, hónap).

**Új termék felvétele:** nyiss egy új sort, töltsd ki legalább a Megnevezést,
a Vonalkódot és a Bruttó eladás árat (a képleteket a meglévő sorból másold le).
Mentés után a programban kattints a **„Készlet újratöltése”** gombra (vagy
indítsd újra), és már beolvasható.

---

## Automatikus frissítés

A program magától észreveszi, ha kiadtál egy újabb változatot, és az **Admin →
⬆ Frissítés** fülön egy gombnyomással telepíthető.

### A boltban

Semmi teendő. Indításkor a program a **háttérben, csendben** megnézi, van-e
újabb verzió – ha nincs net vagy hiba van, egyszerűen nem történik semmi, a
kassza sosem áll meg emiatt. A dolgozó ebből semmit nem lát; az értesítés csak
az admin felületen jelenik meg.

Telepítéskor a program:

1. letölti az új csomagot,
2. **ellenőrzi** (van-e benne épkézláb, hibátlan `keszlet_app.py`),
3. a `backup/` mappába lementi mindazt, amit felül fog írni,
4. kicseréli a program fájljait, és újraindul.

**Amihez soha nem nyúl:** `keszlet.xlsx`, `config.json`, `napi_munkamenet.json`,
`backup/`, és minden `.xlsx`/`.csv` fájl. Ha a másolás közben bármi hiba
történik, automatikusan **visszaáll** az előző állapot a friss mentésből.

### Beállítás (egyszer, neked)

A frissítés **GitHub-kiadásokból** (Releases) dolgozik, és alapból a hivatalos
tárolóra néz:

> https://github.com/Botirocky/oxigen-keszletezo

Ehhez **nem kell semmit beállítani** – egy friss telepítés is azonnal kapja a
frissítéseket. Ha másik tárolót használnál, az Admin → ⬆ Frissítés → **Tároló
beállítása** gombbal adhatod meg (`tulajdonos/tároló` alakban); ez a
`config.json` `update_repo` kulcsába kerül. Ha ezt **üresre** állítod, a
frissítés kikapcsol. Privát tárolóhoz tehetsz GitHub tokent az `update_token`
kulcsba.

Parancssorból is megnézhető:

```bash
python frissito.py              # van-e újabb verzió?
python frissito.py --telepit    # le is tölti és telepíti
```

### Új verzió kiadása (fejlesztői gépen)

```bash
./kiadas.sh --javitas           # 1.1.0 -> 1.1.1  (hibajavítás)
./kiadas.sh --funkcio           # 1.1.0 -> 1.2.0  (új funkció)
./kiadas.sh 1.2.0 -m "Mi változott"
./kiadas.sh 1.2.0 --csak-zip    # csak a csomag, feltöltés nélkül
```

A script felemeli a verziószámot, ellenőrzi hogy a program fordul-e,
összecsomagolja, **kihagyja belőle a bolti adatokat** (`keszlet.xlsx`,
`config.json`), és `gh`-val fel is tölti GitHub-kiadásként. A bolti gép a
következő indításkor meglátja.

## Fontos tudnivalók

- **Egyszerre csak egy helyen írj:** ha a program eladást rögzít, az Excel
  legyen bezárva. Olvasni/eladni közben nyugodtan nyitva lehet, csak a
  rögzítés pillanatában ne.
- A program **biztonságos**: minden rögzítéskor frissen beolvassa a fájlt, így
  a közben Excelben végzett kézi módosításaidat **nem írja felül**.
- **Automatikus biztonsági mentés:** minden rögzítés (eladás, elvitel, új/
  szerkesztett termék) **előtt** a program csendben lement egy másolatot a
  `keszlet.xlsx` felülírás előtti állapotáról a **`backup/`** mappába
  (`keszlet_ÉÉÉÉ-HH-NN_ÓÓPPMM.xlsx`). Ez teljesen a háttérben történik, nem kell
  vele foglalkozni. Ha valaha megsérül vagy elromlik a fájl, egyszerűen másold
  vissza a legfrissebb `backup/`-beli másolatot. **Megőrzés:** a mai nap összes
  mentése + korábbi napokból naponta a legutolsó, **30 napig** visszamenőleg (a
  régebbieket automatikusan törli, hogy ne fogyjon a hely).
- Az `config.json` az Excel elérési útját, a dolgozók névsorát és az admin
  jelszó hash-ét tárolja (a jelszó nem olvasható formában).

---

## Fejlesztői mód (DEV)

Csak teszteléshez/hibakereséshez. **Bekapcsolás háromféleképp:**

- `python keszlet_app.py --dev`
- `KESZLET_DEV=1` környezeti változó
- az appon belül bármikor: **`Ctrl+Shift+D`** (ki is kapcsol)

Bekapcsolva egy indigó **DEV-sáv** jelenik meg a fejléc alatt, ezekkel:

- **Szárazfutás** (alapból BE): a rögzítés **NEM ír az Excelbe**, csak kijelzi
  mit csinálna – így nyugodtan tesztelhetsz a valós adatok kockáztatása nélkül.
- **🎲 Random** / **📋 Lista…**: vonalkódolvasó nélkül is „beolvashatsz” egy
  létező terméket (véletlen, vagy kereshető listából).
- **↩ Visszavonás**: az utolsó **éles** rögzítés visszacsinálása (a készlet
  visszaáll, elvitelnél a Napló-sorok is törlődnek).
- **🧪 Teszt-fájl**: átvált egy `keszlet_teszt.xlsx` másolatra, így az éles
  íráshoz sem a valódi fájlt módosítod.
- **🔎 Diagnosztika**: felismert oszlopok, munkalap, blokk-kezdő oszlopok,
  fájl útja, minta-készletek.

---

## Fájlok

| Fájl / mappa | Szerep |
|------|--------|
| `keszlet_app.py` | maga a program (eladási felület + Excel-kezelés) |
| `keszlet.xlsx`   | az adatbázis (a saját Excel-táblád) |
| `backup/`        | automatikus biztonsági mentések a `keszlet.xlsx`-ről (a program hozza létre) |
| `assets/`        | a programba épített logó-/ikonképek (PNG) |
| `logó/`          | a logó eredeti, nagy felbontású forrásfájljai |
| `inditas.bat`    | Windows indító (telepít + indít) |
| `inditas.sh`     | Linux/macOS indító |
| `epites.bat`     | Windows: önálló `.exe` készítése (mindent letölt + buildel) |
| `epites.sh`      | Linux/macOS: önálló bináris készítése |
| `requirements.txt` | a szükséges csomag (`openpyxl`) |
| `config.json`    | az Excel elérési útja, a dolgozók névsora, az admin jelszó hash-e és a frissítési forrás |
| `frissito.py`    | az automatikus frissítés (GitHub-kiadásokból); önállóan is futtatható |
| `kiadas.sh`      | új verzió kiadása: verzióemelés + csomagolás + GitHub-kiadás |

> **Megjegyzés:** a korábbi, SQLite-alapú változat fájljait (`main.py`,
> `database.py`, `dialogs.py`, `scanner.py`, `inventory.db`) eltávolítottuk –
> ez a program kizárólag az Excelt használja adatbázisként.
