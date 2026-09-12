"""
Frissítő – a program automatikus frissítése GitHub Releases-ből
================================================================

Ez a modul önállóan működik (csak a Python beépített csomagjait használja,
tkinter NEM kell hozzá), így külön is tesztelhető:

    python frissito.py            – megnézi, van-e újabb verzió
    python frissito.py --telepit  – le is tölti és telepíti

Hogyan működik:
  • A `config.json` „update_repo” kulcsában áll a GitHub-tároló, pl.
    "buczibotond/oxigen-keszletezo". Ha ez üres, a frissítés ki van kapcsolva.
  • A modul lekéri a tároló LEGUTÓBBI kiadását (release), és összeveti a
    kiadás verzióját (tag, pl. "v1.2.0") a lenti APP_VERSION-nal.
  • Ha van újabb, letölti a kiadáshoz csatolt .zip-et, kicsomagolja egy ideiglenes
    mappába, ELLENŐRZI (van-e benne épkézláb keszlet_app.py), lement mindent,
    amit felül fog írni a `backup/` mappába, és csak utána másol.
  • Az ADATOKAT soha nem bántja: a keszlet.xlsx, a config.json, a napi
    munkamenet és a backup mappa mindig érintetlen marad (lásd PROTECTED).
  • Ha a másolás közben bármi hiba történik, automatikusan visszaállítja az
    előző állapotot a frissen készült mentésből.

Privát (nem publikus) tároló esetén a `config.json` „update_token” kulcsába
tehető egy GitHub personal access token – ekkor azzal hitelesít.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

# A program verziója. EZ az egyetlen hiteles hely – a kiadas.sh is ezt írja át.
APP_VERSION = "1.1.0"

SCRIPT_DIR = Path(__file__).resolve().parent

if getattr(sys, "frozen", False):
    DATA_DIR = Path(sys.executable).resolve().parent
else:
    DATA_DIR = SCRIPT_DIR

CONFIG_FILE = DATA_DIR / "config.json"
BACKUP_DIR = DATA_DIR / "backup"
UPDATE_LOG = BACKUP_DIR / "frissitesek.log"

API_ROOT = "https://api.github.com"
USER_AGENT = f"OxigenKeszletezo/{APP_VERSION}"
CHECK_TIMEOUT = 8       # másodperc – az indításkori csendes ellenőrzéshez
DOWNLOAD_TIMEOUT = 120   # másodperc – a zip letöltéséhez

# Ezeket a frissítés SOHA nem írja felül és nem törli (adatok és helyi holmik).
PROTECTED_NAMES = {
    "config.json", "napi_munkamenet.json",
    "backup", "logó", "logo", "graphify-out",
    ".venv-build", ".venv", "venv", "build", "dist", "__pycache__", ".git",
}
# Ilyen kiterjesztésű fájl csak akkor másolódik, ha még nem létezik
# (így a bolti keszlet.xlsx biztosan nem íródik felül a sablonnal).
PROTECTED_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".csv"}

# Ha frozen (.exe) módban cserélünk binárist, ide kerül az indítandó cserélő.
_PENDING_SWAP: Path | None = None


class UpdateError(Exception):
    """Felhasználónak megmutatható, magyar nyelvű frissítési hiba."""


class UpdateInfo:
    """Egy elérhető kiadás adatai."""

    def __init__(self, version: str, tag: str, notes: str, url: str,
                 asset_name: str = "", size: int = 0, published: str = "",
                 api_asset: bool = False):
        self.version = version
        self.tag = tag
        self.notes = notes
        self.url = url
        self.asset_name = asset_name
        self.size = size
        self.published = published
        self.api_asset = api_asset      # a letöltés a GitHub API-n megy (token)

    def __repr__(self):  # pragma: no cover – csak hibakereséshez
        return f"<UpdateInfo {self.version} ({self.asset_name or 'zipball'})>"


# ---------------------------------------------------------------- verziók ---

_VER_RE = re.compile(r"\d+")


def parse_version(text) -> tuple[int, ...]:
    """„v1.2.3-beta” → (1, 2, 3). Ami nem szám, azt figyelmen kívül hagyja."""
    if not text:
        return ()
    head = str(text).strip().lstrip("vV").split("-")[0].split("+")[0]
    return tuple(int(n) for n in _VER_RE.findall(head)) or ()


def is_newer(remote, local=APP_VERSION) -> bool:
    """Újabb-e a `remote` verzió a `local`-nál? (hiányzó tagok = 0)"""
    a, b = parse_version(remote), parse_version(local)
    if not a:
        return False
    n = max(len(a), len(b))
    a += (0,) * (n - len(a))
    b += (0,) * (n - len(b))
    return a > b


# ----------------------------------------------------------------- config ---

def _load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _update_config(**changes) -> None:
    """Összeolvasztó mentés – a config.json többi kulcsát nem bántja."""
    cfg = _load_config()
    cfg.update(changes)
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2),
                               encoding="utf-8")
    except OSError:
        pass


def get_repo() -> str:
    """A frissítési forrás („tulajdonos/tároló”), üres sztring = kikapcsolva."""
    return str(_load_config().get("update_repo") or "").strip().strip("/")


def set_repo(repo: str) -> None:
    _update_config(update_repo=str(repo or "").strip().strip("/"))


def get_token() -> str:
    return str(_load_config().get("update_token") or "").strip()


def is_enabled() -> bool:
    return bool(get_repo())


# -------------------------------------------------------------- hálózat -----

def _open(url: str, accept: str, timeout: int):
    headers = {"User-Agent": USER_AGENT, "Accept": accept}
    token = get_token()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, headers=headers)
    return urllib.request.urlopen(req, timeout=timeout)


def _friendly_error(exc: Exception, repo: str) -> UpdateError:
    if isinstance(exc, urllib.error.HTTPError):
        if exc.code == 404:
            return UpdateError(
                f"Nem található kiadás ehhez: {repo}.\n"
                "Ellenőrizd a tároló nevét (Admin → Frissítés), illetve hogy "
                "készült-e már kiadás (release) a GitHubon.\n"
                "Privát tárolóhoz token kell a config.json „update_token” "
                "kulcsában.")
        if exc.code in (401, 403):
            return UpdateError(
                "A GitHub elutasította a kérést (jogosultság vagy óránkénti "
                "kérés-limit). Próbáld később, vagy adj meg tokent.")
        return UpdateError(f"GitHub hiba: {exc.code} {exc.reason}")
    if isinstance(exc, urllib.error.URLError):
        return UpdateError("Nincs internetkapcsolat, vagy a GitHub nem "
                           f"érhető el.\n({exc.reason})")
    return UpdateError(f"Váratlan hiba a frissítés közben:\n{exc}")


def check(repo: str | None = None, timeout: int = CHECK_TIMEOUT) -> UpdateInfo | None:
    """A legutóbbi kiadás lekérése. Újabb verziónál UpdateInfo, különben None.

    Hiba esetén UpdateError-t dob (a csendes változat: `check_quiet`).
    """
    repo = (repo or get_repo()).strip().strip("/")
    if not repo:
        return None
    if repo.count("/") != 1:
        raise UpdateError('A tároló neve „tulajdonos/tároló” alakú legyen, '
                          f'pl. „buczibotond/oxigen-keszletezo”. (Most: „{repo}”)')

    url = f"{API_ROOT}/repos/{repo}/releases/latest"
    try:
        with _open(url, "application/vnd.github+json", timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        raise _friendly_error(exc, repo) from exc

    tag = str(data.get("tag_name") or data.get("name") or "").strip()
    if not tag:
        raise UpdateError("A kiadásnak nincs verziószáma (tag) a GitHubon.")
    if not is_newer(tag):
        return None

    notes = str(data.get("body") or "").strip()
    published = str(data.get("published_at") or "")[:10]
    token = get_token()

    for asset in data.get("assets") or []:
        name = str(asset.get("name") or "")
        if not name.lower().endswith(".zip"):
            continue
        # Privát tárolónál az API-s URL-t kell használni, tokennel.
        if token and asset.get("url"):
            dl, api = str(asset["url"]), True
        else:
            dl, api = str(asset.get("browser_download_url") or ""), False
        if dl:
            return UpdateInfo(tag.lstrip("vV"), tag, notes, dl, name,
                              int(asset.get("size") or 0), published, api)

    zipball = str(data.get("zipball_url") or "")
    if not zipball:
        raise UpdateError("A kiadáshoz nincs letölthető .zip csatolva.")
    return UpdateInfo(tag.lstrip("vV"), tag, notes, zipball,
                      f"{tag}.zip", 0, published, api_asset=True)


def check_quiet(repo: str | None = None,
                timeout: int = CHECK_TIMEOUT) -> UpdateInfo | None:
    """Ugyanaz, de minden hibát elnyel – indításkori, háttérben futó nézéshez.

    A frissítés-ellenőrzés SOHA nem akaszthatja meg a kasszát.
    """
    try:
        return check(repo, timeout)
    except Exception:
        return None


def check_async(callback) -> threading.Thread:
    """Ellenőrzés háttérszálon; a `callback(info)` a végén fut (info lehet None).

    A callback NEM a Tk főszálán fut – csak eltárolni szabad benne az
    eredményt, a felületet a főszálról (pl. `after`) kell frissíteni.
    """
    def worker():
        info = check_quiet()
        try:
            callback(info)
        except Exception:
            pass

    th = threading.Thread(target=worker, name="frissites-ellenorzes",
                          daemon=True)
    th.start()
    return th


# ------------------------------------------------------------- letöltés -----

def download(info: UpdateInfo, progress=None,
             timeout: int = DOWNLOAD_TIMEOUT) -> Path:
    """A kiadás zip-jének letöltése ideiglenes fájlba. Visszaadja az útvonalat."""
    say = progress or (lambda msg: None)
    say("Letöltés…")
    accept = "application/octet-stream" if info.api_asset else "*/*"
    tmp = Path(tempfile.mkdtemp(prefix="oxigen_frissites_")) / (
        info.asset_name or "frissites.zip")
    try:
        with _open(info.url, accept, timeout) as resp, open(tmp, "wb") as out:
            total = info.size or int(resp.headers.get("Content-Length") or 0)
            done = 0
            while True:
                chunk = resp.read(64 * 1024)
                if not chunk:
                    break
                out.write(chunk)
                done += len(chunk)
                if total:
                    say(f"Letöltés… {done * 100 // total}%")
                else:
                    say(f"Letöltés… {done // 1024} kB")
    except Exception as exc:
        shutil.rmtree(tmp.parent, ignore_errors=True)
        raise _friendly_error(exc, get_repo()) from exc
    say("Letöltve.")
    return tmp


# ------------------------------------------------------------ telepítés -----

def _is_protected(rel: Path) -> bool:
    return any(part in PROTECTED_NAMES for part in rel.parts)


def _safe_extract(zf: zipfile.ZipFile, dest: Path) -> None:
    """Kicsomagolás könyvtár-kitörés („../”) ellen védve."""
    dest = dest.resolve()
    for member in zf.infolist():
        target = (dest / member.filename).resolve()
        if not str(target).startswith(str(dest)):
            raise UpdateError("A letöltött csomag gyanús útvonalat tartalmaz, "
                              "a telepítés megszakadt.")
    zf.extractall(dest)


def _find_root(extracted: Path) -> Path:
    """A kicsomagolt fán belül az a mappa, amelyben a keszlet_app.py van."""
    if (extracted / "keszlet_app.py").exists():
        return extracted
    found = sorted(extracted.rglob("keszlet_app.py"),
                   key=lambda p: len(p.parts))
    if found:
        return found[0].parent
    raise UpdateError("A letöltött csomagban nincs keszlet_app.py – "
                      "a telepítés megszakadt (a program érintetlen).")


def _validate(src: Path) -> None:
    """A letöltött program épségének ellenőrzése telepítés ELŐTT."""
    main_py = src / "keszlet_app.py"
    try:
        code = main_py.read_text(encoding="utf-8")
    except OSError as exc:
        raise UpdateError(f"A letöltött keszlet_app.py nem olvasható: {exc}")
    if len(code) < 20_000:
        raise UpdateError("A letöltött keszlet_app.py gyanúsan kicsi – "
                          "a telepítés megszakadt (a program érintetlen).")
    try:
        compile(code, str(main_py), "exec")
    except SyntaxError as exc:
        raise UpdateError("A letöltött program hibás (szintaktikai hiba: "
                          f"{exc.lineno}. sor) – a telepítés megszakadt.")


def _new_files(src: Path):
    """(forrás, cél-relatív) párok, amiket a frissítés ténylegesen másolna."""
    for path in sorted(src.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(src)
        if _is_protected(rel):
            continue
        if rel.suffix.lower() in PROTECTED_SUFFIXES and (DATA_DIR / rel).exists():
            continue        # a meglévő Excelt soha nem írjuk felül
        yield path, rel


def _backup_current(src: Path, say) -> Path:
    """A felülírásra kerülő fájlok mentése a backup/ mappába."""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dest = BACKUP_DIR / f"frissites_{APP_VERSION}_{stamp}"
    saved = 0
    for _, rel in _new_files(src):
        current = DATA_DIR / rel
        if not current.exists():
            continue
        target = dest / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(current, target)
        saved += 1
    say(f"Biztonsági mentés kész ({saved} fájl).")
    return dest


def _restore(backup: Path) -> None:
    """Visszaállítás a frissen készült mentésből (ha a másolás félbeszakadt)."""
    if not backup.exists():
        return
    for path in backup.rglob("*"):
        if not path.is_file():
            continue
        target = DATA_DIR / path.relative_to(backup)
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
        except OSError:
            pass


def _log(line: str) -> None:
    try:
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        with open(UPDATE_LOG, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {line}\n")
    except OSError:
        pass


def _install_source(src: Path, say) -> Path:
    """Forrásból futó (python keszlet_app.py) telepítés."""
    backup = _backup_current(src, say)
    copied = 0
    try:
        for path, rel in _new_files(src):
            target = DATA_DIR / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
            copied += 1
            say(f"Telepítés… ({copied} fájl)")
    except Exception as exc:
        say("Hiba – visszaállítás az előző állapotra…")
        _restore(backup)
        _log(f"SIKERTELEN frissítés, visszaállítva: {exc}")
        raise UpdateError(
            f"A telepítés megszakadt ({exc}).\nAz előző állapot visszaállt a "
            f"mentésből:\n{backup}")
    say(f"Telepítve ({copied} fájl).")
    return backup


def _install_frozen(src: Path, say) -> Path | None:
    """Lefordított (.exe / bináris) verzió cseréje."""
    global _PENDING_SWAP
    exes = [p for p in src.rglob("*") if p.is_file()
            and (p.suffix.lower() == ".exe"
                 or (os.name != "nt" and p.stat().st_mode & 0o111
                     and p.suffix == ""))]
    if not exes:
        raise UpdateError("A letöltött csomagban nincs futtatható program "
                          "(.exe) – a telepítés megszakadt.")
    new_exe = max(exes, key=lambda p: p.stat().st_size)
    current = Path(sys.executable).resolve()

    if os.name == "nt":
        # Windowson a futó .exe nem írható felül: kirakjuk mellé, és egy kis
        # cserélő szkript váltja le, miután a program kilépett.
        staged = current.with_name(current.name + ".new")
        shutil.copy2(new_exe, staged)
        swapper = current.with_name("frissites_csere.bat")
        swapper.write_text(
            "@echo off\r\n"
            "ping -n 3 127.0.0.1 >nul\r\n"
            f'move /y "{staged}" "{current}" >nul\r\n'
            f'start "" "{current}"\r\n'
            'del "%~f0"\r\n',
            encoding="utf-8")
        _PENDING_SWAP = swapper
    else:
        staged = current.with_name(current.name + ".new")
        shutil.copy2(new_exe, staged)
        staged.chmod(0o755)
        os.replace(staged, current)     # futó bináris átnevezéssel cserélhető
    say("Telepítve.")
    return None


def install(zip_path: Path, progress=None) -> Path | None:
    """A letöltött zip telepítése. Visszaadja a biztonsági mentés mappáját."""
    say = progress or (lambda msg: None)
    say("Csomag kibontása…")
    tmp = Path(tempfile.mkdtemp(prefix="oxigen_kibontas_"))
    try:
        try:
            with zipfile.ZipFile(zip_path) as zf:
                _safe_extract(zf, tmp)
        except zipfile.BadZipFile:
            raise UpdateError("A letöltött fájl sérült (nem érvényes zip). "
                              "Próbáld újra.")
        src = _find_root(tmp)
        _validate(src)
        if getattr(sys, "frozen", False):
            backup = _install_frozen(src, say)
        else:
            backup = _install_source(src, say)
        return backup
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def update_now(info: UpdateInfo, progress=None) -> Path | None:
    """Letöltés + telepítés egy lépésben. Visszaadja a mentés mappáját."""
    zip_path = download(info, progress)
    try:
        backup = install(zip_path, progress)
    finally:
        shutil.rmtree(zip_path.parent, ignore_errors=True)
    _log(f"frissítés: {APP_VERSION} -> {info.version} ({info.tag})")
    return backup


def restart() -> None:
    """A program újraindítása a frissített változattal. Nem tér vissza."""
    try:
        if _PENDING_SWAP is not None:
            subprocess.Popen(["cmd", "/c", "start", "", str(_PENDING_SWAP)],
                             cwd=str(DATA_DIR), shell=False)
        elif getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable] + sys.argv[1:], cwd=str(DATA_DIR))
        else:
            subprocess.Popen([sys.executable, str(SCRIPT_DIR / "keszlet_app.py")]
                             + sys.argv[1:], cwd=str(DATA_DIR))
    except Exception:
        pass
    os._exit(0)


# ------------------------------------------------------ parancssori mód -----

def _cli() -> int:
    repo = get_repo()
    print(f"Jelenlegi verzió : {APP_VERSION}")
    print(f"Frissítési forrás: {repo or '(nincs beállítva – config.json / update_repo)'}")
    if not repo:
        return 1
    try:
        info = check(repo, timeout=15)
    except UpdateError as exc:
        print(f"\nHIBA: {exc}")
        return 2
    if info is None:
        print("\nA program naprakész.")
        return 0
    print(f"\nÚj verzió elérhető: {info.version}  ({info.published})")
    if info.notes:
        print("\n" + info.notes.strip()[:1000])
    if "--telepit" not in sys.argv:
        print("\n(Telepítés: python frissito.py --telepit)")
        return 0
    backup = update_now(info, progress=lambda m: print("  " + m))
    print(f"\nKész. Mentés: {backup}")
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
