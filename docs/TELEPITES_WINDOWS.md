# Telepítés Windowsra

Ehhez az útmutatóhoz **nem kell Python**, és rendszergazdai jog sem. A telepítő
mindent feltesz, amire a programnak szüksége van.

## 1. A telepítő letöltése

1. Nyisd meg a kiadások oldalát:
   **https://github.com/Botirocky/oxigen-keszletezo/releases/latest**
2. Az **Assets** részben töltsd le ezt a fájlt:
   `Oxigen_keszletezo_telepito_vX.Y.Z.exe`

> A `..._windows_vX.Y.Z.zip` és a sima `..._vX.Y.Z.zip` a program automatikus
> frissítéséhez kell. Ezeket nem kell letöltened.

## 2. Telepítés

1. Kattints duplán a letöltött `Oxigen_keszletezo_telepito_....exe` fájlra.
2. **„A Windows megvédte a számítógépet” ablak:** a telepítő nincs fizetős
   tanúsítvánnyal aláírva, ezért a Windows ezt a figyelmeztetést mutatja.
   Kattints a **További információ** linkre, majd a **Futtatás mindenképp**
   gombra.
3. Ha kéri, pipáld be az **Ikon az asztalra** lehetőséget, majd kattints a
   **Telepítés** gombra.
4. A végén hagyd bepipálva **A program indítása** lehetőséget, és kattints a
   **Befejezés** gombra.

Hová kerülnek a fájlok:

| Mi | Hol |
|----|-----|
| maga a program | `%LOCALAPPDATA%\Programs\Oxigen keszletezo\` |
| **a bolti adatok** (`keszlet.xlsx`, beállítások, mentések) | **`Dokumentumok\Oxigén készletező\`** |

Az adatmappát a Start menüben is megtalálod: **Oxigén készletező – adatok (Excel)**.

## 3. A készlet-Excel beállítása (első indítás)

A telepítő **nem tartalmaz** készletet, mert a bolti adatok nem kerülnek fel a
GitHubra. Két lehetőséged van:

- **Ezt ajánlom:** az első indítás **előtt** másold a meglévő `keszlet.xlsx`
  fájlt a `Dokumentumok\Oxigén készletező\` mappába. A program magától
  megtalálja.
- Vagy indítsd el a programot. Ha nem talál Excelt, megkérdezi, hol van, és
  megjegyzi a választásodat.

Az Excelt ezután is a megszokott módon szerkesztheted (árak, új termékek).
Közben csak a **program legyen bezárva**, különben a mentésnél szólni fog.

## 4. Első belépés

- **🛒 Dolgozó:** kassza, rögtön használható.
- **🔑 Admin:** az alap jelszó **`admin`**. Az első belépés után **változtasd
  meg** az *Admin → Dolgozók* fülön. Itt veheted fel a dolgozók nevét is.

## 5. Frissítés

Magától megy. A program induláskor csendben megnézi, van-e új verzió. Ha van,
az **Admin → ⬆ Frissítés** fülön egy gombnyomással telepíthető, és a program
újraindul. A bolti adatokhoz (Excel, beállítások, mentések) a frissítés nem nyúl.

Egy újabb telepítőt a régi fölé is feltelepíthetsz. Az adatok ilyenkor is
megmaradnak.

## 6. Átállás a régi, telepítés nélküli `.exe`-ről

Ha eddig az `Oxigen_keszletezo.exe` egy mappából (pl. az Asztalról) futott:

1. Zárd be a programot.
2. A régi mappából másold át a `Dokumentumok\Oxigén készletező\` mappába:
   `keszlet.xlsx`, `config.json` (itt vannak a dolgozók és az admin jelszó)
   és a `backup` mappa.
3. Telepítsd a programot a fenti 1–2. lépés szerint.
4. Ha minden működik, a régi mappa törölhető.

## 7. Eltávolítás

**Gépház → Alkalmazások → Telepített alkalmazások → Oxigén készletező →
Eltávolítás.**

Az eltávolítás **csak a programot** törli. A `Dokumentumok\Oxigén készletező\`
mappa a bolti adatokkal megmarad. Ha azt is törölni akarod, töröld kézzel.

## Gyakori gondok

| Tünet | Megoldás |
|-------|----------|
| „A Windows megvédte a számítógépet” | **További információ → Futtatás mindenképp** (lásd 2. pont). |
| A vírusirtó karanténba tette az exe-t | A program nincs aláírva, ezért egyes vírusirtók gyanúsnak jelölik. Engedélyezd a vírusirtóban, vagy jelentsd téves riasztásként. |
| „Az Excel meg van nyitva” üzenet eladáskor | Zárd be a `keszlet.xlsx`-et az Excelben, majd próbáld újra. |
| Nem találja az Excelt | Tedd a `keszlet.xlsx`-et a `Dokumentumok\Oxigén készletező\` mappába, vagy válaszd ki a program alján az **„Excel kiválasztása…”** gombbal. |
