"""
Készletkezelő – Eladás (egyszerű kasszafelület)
===============================================

Egy nagyon egyszerű, vonalkódolvasóra optimalizált eladási felület, amelynek
az adatbázisa közvetlenül egy Excel-fájl (alapból: keszlet.xlsx).

Hogyan működik:
  • A felület alján mindig egy beolvasó mező van fókuszban.
  • Vonalkódolvasóval (vagy kézzel + Enter) beolvasol egy terméket.
  • A program kikeresi az Excelből, és hozzáadja a "kosárhoz", kiírja a nevét
    és az árát (bruttó eladási ár).
  • Több terméket is be tudsz olvasni egymás után. Ha ugyanazt olvasod be
    kétszer, a darabszám nő.
  • Üres mezőben Entert nyomva (vagy az "Eladás rögzítése" gombbal) megjelenik
    az összeg, és megerősítés után a program az aktuális hónaphoz beírja az
    eladott darabszámot az Excelbe -> így csökken a készlet.

Az Excelt bármikor megnyithatod és szerkesztheted külön (árak, új termékek
felvitele stb.). Fontos: az eladás RÖGZÍTÉSÉHEZ az Excel legyen bezárva, mert
a program akkor tud beleírni a fájlba.

Indítás:  python keszlet_app.py   (vagy: dupla kattintás az inditas.bat-ra)
Függőség: openpyxl  (az inditas.bat ezt automatikusan telepíti)
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import openpyxl
except ImportError:
    import tkinter.messagebox as mb
    _r = tk.Tk(); _r.withdraw()
    mb.showerror(
        "Hiányzó csomag",
        "Az 'openpyxl' csomag nincs telepítve.\n\n"
        "Telepítés parancssorból:\n    pip install openpyxl\n\n"
        "Vagy indítsd a programot az inditas.bat fájllal, az automatikusan telepíti.",
    )
    sys.exit(1)

try:
    import frissito                       # automatikus frissítés (GitHub Releases)
    APP_VERSION = frissito.APP_VERSION
except Exception:                         # a kassza frissítő nélkül is működik
    frissito = None
    APP_VERSION = "1.1.0"

SCRIPT_DIR = Path(__file__).resolve().parent

if getattr(sys, "frozen", False):
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SCRIPT_DIR))
    DATA_DIR = Path(sys.executable).resolve().parent
else:
    RESOURCE_DIR = SCRIPT_DIR
    DATA_DIR = SCRIPT_DIR

CONFIG_FILE = DATA_DIR / "config.json"
SESSION_FILE = DATA_DIR / "napi_munkamenet.json"

BACKUP_DIR = DATA_DIR / "backup"
BACKUP_KEEP_DAYS = 30

SHEET_NAME = "ÖSSZES"

HEADERS = {
    "megnevezés": "name",
    "egyéb": "other",
    "vonalkód": "barcode",
    "forgalmazó": "distributor",
    "áfa": "vat",
    "nettó beszerzés": "net_cost",
    "bruttó eladás": "gross_sale",
    "nyitó készlet": "opening",
    "készlet": "stock_col",
    "készlet érték": "stock_value_col",
}
SALE_BLOCK = "eladás"
PURCHASE_BLOCK = "beszerzés"
WITHDRAW_BLOCK = "elvitel"
LOG_SHEET = "Napló"
SALES_LOG_SHEET = "Eladások"

SALES_LOG_HEADER = ["Dátum", "Eladó", "Vonalkód", "Megnevezés", "Méret",
                    "Darab", "Egységár", "Összeg", "Hónap"]

DEFAULT_ADMIN_PASSWORD = "admin"

DATA_START_ROW = 3

WITHDRAW_REASONS = [
    "Sérült / selejt",
    "Visszáru beszállítónak",
    "Saját / egyéb",
]

HU_MONTHS = [
    "január", "február", "március", "április", "május", "június",
    "július", "augusztus", "szeptember", "október", "november", "december",
]

FONT_UI = "Segoe UI"
FONT_MONO = "Consolas"

_UI_PREFS = ["Segoe UI", "Cantarell", "Ubuntu", "Noto Sans", "DejaVu Sans",
             "Liberation Sans", "Helvetica Neue", "Arial"]
_MONO_PREFS = ["Consolas", "DejaVu Sans Mono", "Ubuntu Mono", "Liberation Mono",
               "Noto Sans Mono", "Menlo", "Courier New"]

def _select_fonts(root) -> None:
    """A FONT_UI / FONT_MONO globálisok beállítása az elérhető betűtípusokra.

    Egy létező Tk-gyökeret igényel (a tkinter.font.families() ezt kéri).
    Ha nem találja egyiket sem a preferáltak közül, a meglévő érték marad,
    amit a Tk maga próbál helyettesíteni.
    """
    global FONT_UI, FONT_MONO
    try:
        import tkinter.font as tkfont
        available = set(tkfont.families(root))
    except Exception:
        return
    for fam in _UI_PREFS:
        if fam in available:
            FONT_UI = fam
            break
    for fam in _MONO_PREFS:
        if fam in available:
            FONT_MONO = fam
            break

def format_ft(value: float | int | None) -> str:
    """Magyaros pénzformátum, pl. 27990 -> '27 990 Ft'."""
    try:
        n = int(round(float(value or 0)))
    except (TypeError, ValueError):
        n = 0
    return f"{n:,}".replace(",", " ") + " Ft"

def as_text(value) -> str:
    """Cellaérték szöveggé alakítása felesleges '.0' nélkül."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def parse_price(value) -> float | None:
    """Felhasználói árbevitel számmá alakítása.

    Megengedő: kezeli a szóközt, nem törhető szóközt, a 'Ft' utótagot és a
    tizedesvesszőt is (pl. '27 990 Ft', '2799,5'). Üres/érvénytelen -> None.
    """
    s = as_text(value).lower().replace("ft", "")
    for ch in (" ", "\t", " "):
        s = s.replace(ch, "")
    s = s.replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None

def parse_int(value) -> int:
    """Felhasználói darabszám-bevitel egész számmá alakítása (üres/hibás -> 0)."""
    s = as_text(value)
    for ch in (" ", "\t", " "):
        s = s.replace(ch, "")
    s = s.replace(",", ".")
    if not s:
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0

def norm_barcode(value) -> str:
    """Vonalkód egységes, kanonikus kulcsa.

    Ugyanaz a kód sokféleképp érkezhet: az Excelben lehet szám vagy szöveg, a
    vonalkódolvasó pedig küldhet köré tabot, újsort, normál vagy nem törhető
    szóközt, sőt egy-egy cellából lebegőpontos '.0' farok is jöhet. Ez mindezt
    egyetlen, szóköz nélküli alakra hozza, hogy a beolvasás és a fájlba írás
    biztosan ugyanazt a terméket találja meg.

    A vezető nullákat MEGTARTJA – a nulla-toleráns visszaesés a kereső
    (`Inventory._match`) feladata, hogy a szándékosan eltérő kódok ne keveredjenek.
    """
    s = as_text(value)
    if not s:
        return ""
    for ch in (" ", "\t", "\r", "\n", " "):
        s = s.replace(ch, "")
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s

def _blend(c1: str, c2: str, t: float) -> str:
    """Két #RRGGBB szín lineáris keverése (t=0 -> c1, t=1 -> c2)."""
    t = max(0.0, min(1.0, t))
    a = tuple(int(c1[i:i + 2], 16) for i in (1, 3, 5))
    b = tuple(int(c2[i:i + 2], 16) for i in (1, 3, 5))
    m = tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return f"#{m[0]:02x}{m[1]:02x}{m[2]:02x}"

class Product:
    """Egy termék (az Excel egy sora) memóriában tartott pillanatképe."""

    __slots__ = ("row", "name", "other", "barcode", "distributor",
                 "vat", "net_cost", "gross_sale", "opening", "stock")

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    @property
    def label(self) -> str:
        """Megjelenítendő név, pl. '73690 / 38'."""
        if self.other:
            return f"{self.name} / {self.other}"
        return str(self.name)

class Inventory:
    """Az Excel-fájl beolvasása és az eladások visszaírása."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.products: list[Product] = []
        self.by_barcode: dict[str, Product] = {}
        self.by_barcode_zs: dict[str, Product] = {}
        self.cols: dict[str, int] = {}
        self.sale_start_col: int = 0
        self.purchase_start_col: int = 0
        self.withdraw_start_col: int = 0
        self.sheet_name: str = SHEET_NAME
        self.last_log_stamp: str | None = None
        self.load()

    def _save_workbook(self, wb) -> None:
        """A munkafüzet mentése – de ELŐTTE csendben lementi a felülírás előtti,
        lemezen lévő állapotot. Minden éles Excel-írás ezen megy át, így a
        biztonsági mentés egy helyen, automatikusan megtörténik."""
        self._backup_workbook()
        wb.save(self.path)

    def _backup_workbook(self) -> None:
        """Csendes biztonsági másolat a jelenlegi (még felül nem írt) Excelről.
        A mentés hibája SOSEM akadályozhatja meg a rögzítést, ezért mindent
        némán elnyel."""
        try:
            src = self.path
            if not src.exists():
                return
            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            now = datetime.now()
            dst = BACKUP_DIR / f"keszlet_{now:%Y-%m-%d_%H%M%S}.xlsx"
            if not dst.exists():
                shutil.copy2(src, dst)
            self._prune_backups()
        except Exception:
            pass

    @staticmethod
    def _prune_backups() -> None:
        """Retenció: a mai összes mentés + korábbi napokból naponta a legutolsó
        marad, BACKUP_KEEP_DAYS napig visszamenőleg. A régebbieket törli."""
        try:
            files = sorted(BACKUP_DIR.glob("keszlet_*.xlsx"))
        except Exception:
            return
        today = datetime.now().date()
        keep: set[Path] = set()
        last_of_day: dict = {}
        for f in files:
            stamp = f.stem[len("keszlet_"):]
            try:
                d = datetime.strptime(stamp, "%Y-%m-%d_%H%M%S").date()
            except ValueError:
                keep.add(f)
                continue
            if (today - d).days > BACKUP_KEEP_DAYS:
                continue
            if d == today:
                keep.add(f)
            else:
                last_of_day[d] = f
        keep.update(last_of_day.values())
        for f in files:
            if f not in keep:
                try:
                    f.unlink()
                except OSError:
                    pass

    def _detect_columns(self, ws) -> None:
        self.cols = {}
        self.sale_start_col = 0
        self.purchase_start_col = 0
        self.withdraw_start_col = 0
        for col in range(1, ws.max_column + 1):
            raw = ws.cell(row=1, column=col).value
            if raw is None:
                continue
            text = str(raw).strip().lower()
            if text in HEADERS:
                self.cols[HEADERS[text]] = col
            if text == SALE_BLOCK:
                self.sale_start_col = col
            elif text == PURCHASE_BLOCK:
                self.purchase_start_col = col
            elif text == WITHDRAW_BLOCK:
                self.withdraw_start_col = col

        self.cols.setdefault("name", 1)
        self.cols.setdefault("other", 2)
        self.cols.setdefault("barcode", 3)
        self.cols.setdefault("distributor", 4)
        self.cols.setdefault("vat", 5)
        self.cols.setdefault("net_cost", 6)
        self.cols.setdefault("gross_sale", 8)
        self.cols.setdefault("opening", 9)
        self.cols.setdefault("stock_col", 36)
        self.cols.setdefault("stock_value_col", 37)
        if not self.purchase_start_col:
            self.purchase_start_col = 10
        if not self.sale_start_col:
            self.sale_start_col = 23

    def _pick_sheet(self, wb):
        if SHEET_NAME in wb.sheetnames:
            return wb[SHEET_NAME]
        for ws in wb.worksheets:
            for col in range(1, min(ws.max_column, 50) + 1):
                v = ws.cell(row=1, column=col).value
                if v and str(v).strip().lower() == "vonalkód":
                    return ws
        return wb.worksheets[0]

    def load(self) -> None:
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            self.sheet_name = ws.title
            self._detect_columns(ws)

            products: list[Product] = []
            by_barcode: dict[str, Product] = {}
            by_barcode_zs: dict[str, Product] = {}
            c = self.cols
            for row in range(DATA_START_ROW, ws.max_row + 1):
                name = ws.cell(row=row, column=c["name"]).value
                barcode = norm_barcode(ws.cell(row=row, column=c["barcode"]).value)
                if not name and not barcode:
                    continue

                purchases = self._sum_block(ws, row, self.purchase_start_col)
                sales = self._sum_block(ws, row, self.sale_start_col)
                withdrawals = (self._sum_block(ws, row, self.withdraw_start_col)
                               if self.withdraw_start_col else 0.0)
                opening = self._num(ws.cell(row=row, column=c["opening"]).value)

                p = Product(
                    row=row,
                    name=as_text(name),
                    other=as_text(ws.cell(row=row, column=c["other"]).value),
                    barcode=barcode,
                    distributor=as_text(ws.cell(row=row, column=c["distributor"]).value),
                    vat=self._num(ws.cell(row=row, column=c["vat"]).value),
                    net_cost=self._num(ws.cell(row=row, column=c["net_cost"]).value),
                    gross_sale=self._num(ws.cell(row=row, column=c["gross_sale"]).value),
                    opening=opening,
                    stock=opening + purchases - sales - withdrawals,
                )
                products.append(p)
                if barcode:
                    by_barcode[barcode] = p
                    zs = barcode.lstrip("0")
                    if zs and zs != barcode:
                        by_barcode_zs.setdefault(zs, p)

            self.products = products
            self.by_barcode = by_barcode
            self.by_barcode_zs = by_barcode_zs
        finally:
            wb.close()

    @staticmethod
    def _num(value) -> float:
        """Számértékké alakítás; képletet/szöveget 0-nak vesz."""
        if isinstance(value, (int, float)):
            return float(value)
        return 0.0

    def _sum_block(self, ws, row: int, start_col: int) -> float:
        total = 0.0
        for col in range(start_col, start_col + 12):
            total += self._num(ws.cell(row=row, column=col).value)
        return total

    def find(self, code: str) -> Product | None:
        return self._match(code, self.by_barcode, self.by_barcode_zs)

    @staticmethod
    def _match(code: str, primary: dict, secondary: dict):
        """Kanonikus, majd nulla-toleráns keresés a megadott indexekben.

        Először pontos (kanonikus) egyezést keres. Ha nincs, a vezető nullák
        elhagyásával mindkét irányban próbálkozik – így a '0048' beolvasás a
        '48' termékre talál és fordítva. Találat híján None.
        """
        key = norm_barcode(code)
        if not key:
            return None
        hit = primary.get(key)
        if hit is not None:
            return hit
        zs = key.lstrip("0") or "0"
        if zs != key:
            hit = primary.get(zs)
            if hit is not None:
                return hit
        return secondary.get(zs)

    @staticmethod
    def _index_rows(ws, bc_col: int) -> tuple[dict[str, int], dict[str, int]]:
        """A lap vonalkód-oszlopából {kanonikus -> sor} és {nulla nélkül -> sor}
        indexeket épít – ugyanazzal a logikával, mint a betöltés."""
        primary: dict[str, int] = {}
        secondary: dict[str, int] = {}
        for row in range(DATA_START_ROW, ws.max_row + 1):
            key = norm_barcode(ws.cell(row=row, column=bc_col).value)
            if not key:
                continue
            primary[key] = row
            zs = key.lstrip("0")
            if zs and zs != key:
                secondary.setdefault(zs, row)
        return primary, secondary

    def commit_sale(self, items: list[tuple[str, int]], month: int,
                    seller: str | None = None) -> list[str]:
        """Eladás rögzítése: az adott hónap eladás-oszlopához hozzáadja a db-ot.

        `items`: (vonalkód, darabszám) párok listája.
        `month`: 1-12.
        `seller`: ha meg van adva (éles eladás), minden tételt naplóz az
            'Eladások' lapra (dátum, eladó, db, egységár, összeg) – így az
            admin később lekérdezheti, ki mit adott el. A DEV-visszavonás
            seller nélkül hív, ezért nem naplóz.
        Visszaadja a NEM megtalált vonalkódok listáját (általában üres).
        A fájlt frissen nyitja meg, hogy a közben Excelben végzett kézi
        módosításokat ne írja felül.
        """
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            self._detect_columns(ws)
            bc_col = self.cols["barcode"]
            name_col = self.cols["name"]
            other_col = self.cols["other"]
            price_col = self.cols["gross_sale"]
            sale_col = self.sale_start_col + (month - 1)

            row_of, row_of_zs = self._index_rows(ws, bc_col)

            log_ws = None
            stamp = ""
            if seller is not None:
                self._ensure_sales_log(wb)
                log_ws = wb[SALES_LOG_SHEET]
                stamp = datetime.now().strftime("%Y.%m.%d %H:%M:%S")

            missing: list[str] = []
            for barcode, qty in items:
                row = self._match(barcode, row_of, row_of_zs)
                if row is None:
                    missing.append(barcode)
                    continue
                cell = ws.cell(row=row, column=sale_col)
                current = cell.value if isinstance(cell.value, (int, float)) else 0
                cell.value = int(current) + int(qty)

                if log_ws is not None and int(qty) > 0:
                    name = as_text(ws.cell(row=row, column=name_col).value)
                    size = as_text(ws.cell(row=row, column=other_col).value)
                    price = self._num(ws.cell(row=row, column=price_col).value)
                    log_ws.append([stamp, seller, barcode, name, size,
                                   int(qty), int(price), int(price) * int(qty),
                                   month])

            self._save_workbook(wb)
        finally:
            wb.close()

        if seller is not None:
            self.last_log_stamp = stamp
        self.load()
        return missing

    @staticmethod
    def _ensure_sales_log(wb):
        """Létrehozza az 'Eladások' munkalapot fejléccel, ha még nincs."""
        if SALES_LOG_SHEET not in wb.sheetnames:
            sh = wb.create_sheet(SALES_LOG_SHEET)
            sh.append(SALES_LOG_HEADER)

    def update_product(self, *, row: int, name: str, other: str, barcode: str,
                       gross_sale: float, opening: int | None = None) -> Product | None:
        """Meglévő termék mezőinek módosítása az adott sorban, majd újratöltés.

        A havi beszerzés/eladás/elvitel oszlopokat NEM érinti, csak a törzs-
        adatokat (név, méret, vonalkód, eladási ár, és ha megadod, nyitókészlet).
        """
        barcode = norm_barcode(barcode)
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            self._detect_columns(ws)
            c = self.cols
            ws.cell(row=row, column=c["name"]).value = name
            ws.cell(row=row, column=c["other"]).value = other or None
            ws.cell(row=row, column=c["barcode"]).value = barcode
            ws.cell(row=row, column=c["gross_sale"]).value = gross_sale
            if opening is not None:
                ws.cell(row=row, column=c["opening"]).value = int(opening)
            self._save_workbook(wb)
        finally:
            wb.close()
        self.load()
        return self.find(barcode)

    def delete_product(self, *, row: int) -> None:
        """Termék törlése: az adott sor TELJES kiürítése, majd újratöltés.

        Szándékosan nem sortörlés (ws.delete_rows): az openpyxl a törölt sor
        alatti sorok képlethivatkozásait nem igazítaná át, így az összes
        lentebbi KÉSZLET-képlet elromlana. Az üres sort a betöltés átugorja,
        és egy későbbi új termék újra felhasználhatja (ha ez az utolsó sor).
        Mentés előtt automatikus biztonsági másolat készül."""
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).value = None
            self._save_workbook(wb)
        finally:
            wb.close()
        self.load()

    def read_log_rows(self) -> list[list]:
        """A 'Napló' (elvitel) lap sorai fejléc nélkül; üres lista, ha nincs."""
        return self._read_sheet_rows(LOG_SHEET)

    def read_sales_log_rows(self) -> list[list]:
        """Az 'Eladások' napló sorai fejléc nélkül; üres lista, ha nincs."""
        return self._read_sheet_rows(SALES_LOG_SHEET)

    def _read_sheet_rows(self, sheet_name: str) -> list[list]:
        try:
            wb = openpyxl.load_workbook(self.path, data_only=True)
        except Exception:
            return []
        try:
            if sheet_name not in wb.sheetnames:
                return []
            ws = wb[sheet_name]
            rows: list[list] = []
            for r in range(2, ws.max_row + 1):
                vals = [ws.cell(row=r, column=col).value
                        for col in range(1, ws.max_column + 1)]
                if any(v not in (None, "") for v in vals):
                    rows.append(vals)
            return rows
        finally:
            wb.close()

    def _last_data_row(self, ws) -> int:
        """Az utolsó olyan sor száma, ahol van név vagy vonalkód (adat).

        Ha nincs egyetlen adatsor sem, a fejléc-blokk utáni sort jelez vissza
        (DATA_START_ROW - 1), így az új sor a DATA_START_ROW lesz.
        """
        c = self.cols
        last = DATA_START_ROW - 1
        for row in range(DATA_START_ROW, ws.max_row + 1):
            has_name = ws.cell(row=row, column=c["name"]).value not in (None, "")
            has_bc = ws.cell(row=row, column=c["barcode"]).value not in (None, "")
            if has_name or has_bc:
                last = row
        return last

    def _write_stock_formula(self, ws, row: int) -> None:
        """Beírja a KÉSZLET-oszlop alapképletét (nyitó + beszerzés − eladás
        [− elvitel]) az adott sorba – ha nincs honnan átmásolni."""
        from openpyxl.utils import get_column_letter
        c = self.cols
        op = get_column_letter(c["opening"])
        ps = get_column_letter(self.purchase_start_col)
        pe = get_column_letter(self.purchase_start_col + 11)
        ss = get_column_letter(self.sale_start_col)
        se = get_column_letter(self.sale_start_col + 11)
        f = (f"={op}{row}+SUM({ps}{row}:{pe}{row})"
             f"-SUM({ss}{row}:{se}{row})")
        if self.withdraw_start_col:
            elv_s = get_column_letter(self.withdraw_start_col)
            elv_e = get_column_letter(self.withdraw_start_col + 11)
            f += f"-SUM({elv_s}{row}:{elv_e}{row})"
        ws.cell(row=row, column=c["stock_col"]).value = f

    def _ensure_sum_formulas(self, ws, row: int) -> bool:
        """A beszerzés/eladás (és elvitel) blokk Σ-cellájába =SUM(...) képletet
        ír, ha még nincs ott – a sorok KÉSZLET-képlete ezekre hivatkozik.

        Csak akkor ír, ha az adott oszlop fejléce tényleg 'Σ'. Visszaadja,
        hogy minden blokk Σ-képlete rendben van-e.
        """
        from openpyxl.utils import get_column_letter
        blocks = [self.purchase_start_col, self.sale_start_col]
        if self.withdraw_start_col:
            blocks.append(self.withdraw_start_col)
        all_ok = True
        for start in blocks:
            col = start + 12
            cell = ws.cell(row=row, column=col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            header = str(ws.cell(row=1, column=col).value or "").strip().lower()
            if header != "σ":
                all_ok = False
                continue
            s = get_column_letter(start)
            e = get_column_letter(start + 11)
            cell.value = f"=SUM({s}{row}:{e}{row})"
        return all_ok

    def _copy_row_formulas(self, ws, src_row: int, dst_row: int) -> None:
        """A számolt oszlopok (KÉSZLET, KÉSZLET ÉRTÉK, elvitel-Σ) képleteit
        átmásolja egy meglévő sorból az új sorba, a sorhivatkozásokat igazítva.

        Így bármilyen képletet kezel, amit a felhasználó az Excelben használ.
        """
        from openpyxl.formula.translate import Translator
        from openpyxl.utils import get_column_letter
        if src_row < DATA_START_ROW:
            return
        cols = [self.cols.get("stock_col"), self.cols.get("stock_value_col"),
                self.purchase_start_col + 12,
                self.sale_start_col + 12]
        if self.withdraw_start_col:
            cols.append(self.withdraw_start_col + 12)
        for col in cols:
            if not col:
                continue
            src = ws.cell(row=src_row, column=col).value
            if isinstance(src, str) and src.startswith("="):
                origin = f"{get_column_letter(col)}{src_row}"
                dest = f"{get_column_letter(col)}{dst_row}"
                ws.cell(row=dst_row, column=col).value = (
                    Translator(src, origin=origin).translate_formula(dest))

    def add_product(self, *, name: str, other: str, barcode: str,
                    gross_sale: float, opening: int = 0, purchase: int = 0,
                    month: int | None = None) -> Product | None:
        """Új termék felvétele: új sort ír az Excelbe, majd újratölt.

        Az Excel logikáját követi: a `purchase` darabszám az adott (alapból az
        aktuális) hónap beszerzés-oszlopába kerül, az `opening` pedig a Nyitó
        készletbe (évkezdő állomány – új terméknél jellemzően 0).

        A KÉSZLET (és ha van, KÉSZLET ÉRTÉK / elvitel-Σ) képleteit a legutolsó
        meglévő adatsorból másolja át; emellett gondoskodik a beszerzés/eladás
        Σ-képleteiről is, mert a KÉSZLET-képlet ezekre hivatkozik. Ha nincs
        honnan másolni, alap KÉSZLET-képletet ír.
        Visszaadja a frissen betöltött Product-ot (vagy None, ha valamiért nem
        található). A fájlt frissen nyitja meg, hogy a kézi módosításokat ne
        írja felül.
        """
        barcode = norm_barcode(barcode)
        if month is None:
            month = datetime.now().month
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            self._detect_columns(ws)
            c = self.cols
            src = self._last_data_row(ws)
            row = src + 1

            ws.cell(row=row, column=c["name"]).value = name
            if other:
                ws.cell(row=row, column=c["other"]).value = other
            ws.cell(row=row, column=c["barcode"]).value = barcode
            ws.cell(row=row, column=c["gross_sale"]).value = gross_sale
            ws.cell(row=row, column=c["opening"]).value = int(opening)
            if int(purchase):
                col = self.purchase_start_col + (int(month) - 1)
                ws.cell(row=row, column=col).value = int(purchase)

            self._copy_row_formulas(ws, src, row)
            sums_ok = self._ensure_sum_formulas(ws, row)
            if (not sums_ok
                    or not isinstance(ws.cell(row=row,
                                              column=c["stock_col"]).value, str)):
                self._write_stock_formula(ws, row)

            self._save_workbook(wb)
        finally:
            wb.close()

        self.load()
        return self.find(barcode)

    def compute_stats(self, low_threshold: int = 3) -> dict:
        """Összesített statisztika az admin felülethez.

        A havi eladott darabokat közvetlenül az ÖSSZES lap eladás-blokkjából
        olvassa (nyers számok), a bevételt a bruttó eladási árral szorozva.
        A készletet és a készletértéket a memóriabeli (load() által számolt)
        adatokból veszi, így Excel-újraszámolás nélkül is pontos.
        """
        monthly_units = [0] * 12
        monthly_revenue = [0] * 12
        per_product_sold: dict[int, int] = {}

        try:
            wb = openpyxl.load_workbook(self.path, data_only=False)
        except Exception:
            wb = None
        if wb is not None:
            try:
                ws = self._pick_sheet(wb)
                self._detect_columns(ws)
                price_col = self.cols["gross_sale"]
                for row in range(DATA_START_ROW, ws.max_row + 1):
                    price = self._num(ws.cell(row=row, column=price_col).value)
                    sold_here = 0
                    for m in range(12):
                        u = self._num(ws.cell(row=row,
                                              column=self.sale_start_col + m).value)
                        ui = int(u)
                        if ui:
                            monthly_units[m] += ui
                            monthly_revenue[m] += int(price) * ui
                            sold_here += ui
                    if sold_here:
                        per_product_sold[row] = sold_here
            finally:
                wb.close()

        total_units = 0
        total_value = 0
        low_stock: list[Product] = []
        sold_by_row = per_product_sold
        for p in self.products:
            st = int(p.stock or 0)
            total_units += st
            total_value += int((p.stock or 0) * (p.net_cost or 0))
            if st <= low_threshold:
                low_stock.append(p)

        row_to_product = {p.row: p for p in self.products}
        top = []
        for row, sold in sold_by_row.items():
            p = row_to_product.get(row)
            if p is not None:
                top.append((p, sold, int(sold * (p.gross_sale or 0))))
        top.sort(key=lambda t: t[1], reverse=True)
        low_stock.sort(key=lambda p: int(p.stock or 0))

        return {
            "monthly_units": monthly_units,
            "monthly_revenue": monthly_revenue,
            "total_units": total_units,
            "total_value": total_value,
            "year_units": sum(monthly_units),
            "year_revenue": sum(monthly_revenue),
            "top_products": top,
            "low_stock": low_stock,
            "product_count": len(self.products),
            "low_threshold": low_threshold,
        }

    def report_days(self) -> list[str]:
        """Azok a napok ('ÉÉÉÉ.HH.NN'), amelyeken volt eladás vagy elvitel.

        Legújabb nap elöl – a napi zárás nap-választójához.
        """
        days: set[str] = set()
        for r in self.read_sales_log_rows():
            d = as_text(r[0] if r else "")[:10]
            if d:
                days.add(d)
        for r in self.read_log_rows():
            d = as_text(r[0] if r else "")[:10]
            if d:
                days.add(d)
        return sorted(days, reverse=True)

    def daily_summary(self, day: str) -> dict:
        """Egy nap ('ÉÉÉÉ.HH.NN') zárása: eladások és elvitelek összesítése.

        Az 'Eladások' és 'Napló' lapok dátumbélyege 'ÉÉÉÉ.HH.NN ÓÓ:PP', tehát
        az első 10 karakter a nap. A nyugták (tranzakciók) számát a különböző
        időbélyegek adják – egy rögzítés minden tétele azonos bélyeget kap.
        """
        sales = [r for r in self.read_sales_log_rows()
                 if as_text(r[0] if r else "")[:10] == day]
        withdrawals = [r for r in self.read_log_rows()
                       if as_text(r[0] if r else "")[:10] == day]

        total_units = 0
        total_rev = 0
        receipts: set[str] = set()
        per_seller: dict[str, list] = {}
        per_product: dict[str, list] = {}

        for r in sales:
            stamp = as_text(r[0])
            seller = as_text(r[1]) or "—"
            barcode = as_text(r[2])
            name = as_text(r[3])
            size = as_text(r[4])
            qty = parse_int(r[5])
            amount = int(self._num(r[7]))
            total_units += qty
            total_rev += amount
            receipts.add(stamp)
            s = per_seller.setdefault(seller, [0, 0])
            s[0] += qty
            s[1] += amount
            key = barcode or f"{name}|{size}"
            p = per_product.setdefault(key, [name, size, 0, 0])
            p[2] += qty
            p[3] += amount

        seller_rows = sorted(([k, v[0], v[1]] for k, v in per_seller.items()),
                             key=lambda x: x[2], reverse=True)
        product_rows = sorted(per_product.values(),
                              key=lambda x: x[2], reverse=True)
        n = len(receipts)

        return {
            "day": day,
            "total_units": total_units,
            "total_revenue": total_rev,
            "receipts": n,
            "avg_basket": int(total_rev / n) if n else 0,
            "sellers": seller_rows,
            "products": product_rows,
            "withdrawals": withdrawals,
            "withdraw_units": sum(parse_int(r[4]) for r in withdrawals),
        }

    def _ensure_withdraw_structure(self, ws, wb) -> int:
        """Gondoskodik róla, hogy legyen 'elvitel' havi blokk és 'Napló' lap.

        Ha még nincs, létrehozza: az ÖSSZES lapra a KÉSZLET ÉRTÉK után beszúr
        egy 12 hónapos 'elvitel' blokkot + Σ oszlopot, és minden termék-sor
        KÉSZLET-képletét kiegészíti az elvitel levonásával. Visszaadja az
        elvitel-blokk kezdő oszlopát.
        """
        from openpyxl.utils import get_column_letter

        self._detect_columns(ws)
        if self.withdraw_start_col:
            self._ensure_log_sheet(wb)
            return self.withdraw_start_col

        start = max(self.cols.get("stock_value_col", 37),
                    self.cols.get("stock_col", 36)) + 1
        sigma = start + 12

        ws.cell(row=1, column=start).value = WITHDRAW_BLOCK
        for i in range(12):
            ws.cell(row=2, column=start + i).value = f"{i + 1}."
        ws.cell(row=1, column=sigma).value = "Σ"

        elv_s = get_column_letter(start)
        elv_e = get_column_letter(start + 11)
        op = get_column_letter(self.cols["opening"])
        ps = get_column_letter(self.purchase_start_col)
        pe = get_column_letter(self.purchase_start_col + 11)
        ss = get_column_letter(self.sale_start_col)
        se = get_column_letter(self.sale_start_col + 11)
        stock_col = self.cols["stock_col"]
        bc_col = self.cols["barcode"]
        name_col = self.cols["name"]

        for row in range(DATA_START_ROW, ws.max_row + 1):
            has_name = ws.cell(row=row, column=name_col).value not in (None, "")
            has_bc = ws.cell(row=row, column=bc_col).value not in (None, "")
            if not has_name and not has_bc:
                continue
            ws.cell(row=row, column=sigma).value = f"=SUM({elv_s}{row}:{elv_e}{row})"
            cur = ws.cell(row=row, column=stock_col).value
            if isinstance(cur, str) and cur.startswith("="):
                ws.cell(row=row, column=stock_col).value = (
                    f"{cur}-SUM({elv_s}{row}:{elv_e}{row})")
            else:
                ws.cell(row=row, column=stock_col).value = (
                    f"={op}{row}+SUM({ps}{row}:{pe}{row})"
                    f"-SUM({ss}{row}:{se}{row})-SUM({elv_s}{row}:{elv_e}{row})")

        self._ensure_log_sheet(wb)
        self.withdraw_start_col = start
        return start

    @staticmethod
    def _ensure_log_sheet(wb):
        """Létrehozza a 'Napló' munkalapot fejléccel, ha még nincs."""
        if LOG_SHEET not in wb.sheetnames:
            nap = wb.create_sheet(LOG_SHEET)
            nap.append(["Dátum", "Vonalkód", "Megnevezés", "Méret",
                        "Darab", "Indok", "Hónap"])

    def commit_withdrawal(self, items: list[tuple[str, int]], month: int,
                          reason: str) -> list[str]:
        """Elvitel rögzítése: az adott hónap elvitel-oszlopához adja a db-ot,
        és minden tételt naplóz a 'Napló' lapon (dátummal és indokkal).

        Visszaadja a meg nem talált vonalkódok listáját.
        """
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            start = self._ensure_withdraw_structure(ws, wb)
            bc_col = self.cols["barcode"]
            name_col = self.cols["name"]
            other_col = self.cols["other"]
            w_col = start + (month - 1)
            nap = wb[LOG_SHEET]
            stamp = datetime.now().strftime("%Y.%m.%d %H:%M:%S")

            row_of, row_of_zs = self._index_rows(ws, bc_col)

            missing: list[str] = []
            for barcode, qty in items:
                row = self._match(barcode, row_of, row_of_zs)
                if row is None:
                    missing.append(barcode)
                    continue
                cell = ws.cell(row=row, column=w_col)
                current = cell.value if isinstance(cell.value, (int, float)) else 0
                cell.value = int(current) + int(qty)
                name = as_text(ws.cell(row=row, column=name_col).value)
                size = as_text(ws.cell(row=row, column=other_col).value)
                nap.append([stamp, barcode, name, size, int(qty), reason, month])

            self._save_workbook(wb)
        finally:
            wb.close()

        self.last_log_stamp = stamp
        self.load()
        return missing

    def undo_sale(self, items: list[tuple[str, int]], month: int) -> list[str]:
        """Egy eladás visszavonása: a hónap eladás-oszlopából LEVONJA a db-ot.
        (A commit_sale fordítottja – negatív darabszámmal hívja.)"""
        return self.commit_sale([(bc, -int(qty)) for bc, qty in items], month)

    def undo_withdrawal(self, items: list[tuple[str, int]], month: int,
                        stamp: str | None = None) -> list[str]:
        """Egy elvitel visszavonása: a hónap elvitel-oszlopából levonja a db-ot,
        és törli a hozzá tartozó Napló-sorokat (a megadott naplóbélyeg alapján)."""
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            ws = self._pick_sheet(wb)
            self._detect_columns(ws)
            if not self.withdraw_start_col:
                return [bc for bc, _ in items]
            bc_col = self.cols["barcode"]
            w_col = self.withdraw_start_col + (month - 1)
            row_of, row_of_zs = self._index_rows(ws, bc_col)

            missing: list[str] = []
            for barcode, qty in items:
                row = self._match(barcode, row_of, row_of_zs)
                if row is None:
                    missing.append(barcode)
                    continue
                cell = ws.cell(row=row, column=w_col)
                current = cell.value if isinstance(cell.value, (int, float)) else 0
                cell.value = int(current) - int(qty)

            if stamp and LOG_SHEET in wb.sheetnames:
                nap = wb[LOG_SHEET]
                for r in range(nap.max_row, 1, -1):
                    if as_text(nap.cell(row=r, column=1).value) == stamp:
                        nap.delete_rows(r, 1)

            self._save_workbook(wb)
        finally:
            wb.close()

        self.load()
        return missing

def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}

def save_config(cfg: dict) -> None:
    try:
        CONFIG_FILE.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass

def update_config(**changes) -> dict:
    """A config.json összeolvasztó frissítése: csak a megadott kulcsokat írja
    felül, a többit (pl. dolgozók, admin-jelszó) érintetlenül hagyja."""
    cfg = load_config()
    cfg.update(changes)
    save_config(cfg)
    return cfg

def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def get_admin_hash() -> str:
    """Az érvényben lévő admin-jelszó hash-e (alapból a DEFAULT_ADMIN_PASSWORD)."""
    cfg = load_config()
    return cfg.get("admin_password_hash") or _hash_pw(DEFAULT_ADMIN_PASSWORD)

def check_admin_password(password: str) -> bool:
    return _hash_pw(password) == get_admin_hash()

def set_admin_password(password: str) -> None:
    update_config(admin_password_hash=_hash_pw(password))

def get_employees() -> list[str]:
    """A dolgozók (eladók) névsora a config.json-ból."""
    cfg = load_config()
    names = cfg.get("employees")
    if isinstance(names, list):
        return [str(n) for n in names if str(n).strip()]
    return []

def set_employees(names: list[str]) -> None:
    clean: list[str] = []
    for n in names:
        n = str(n).strip()
        if n and n not in clean:
            clean.append(n)
    update_config(employees=clean)

# ---------------------------------------------------------------- Frissítés --
# Az indításkori ellenőrzés a háttérben fut, és SOHA nem állítja meg a kasszát:
# ha nincs net vagy hiba van, egyszerűen nem történik semmi. Az eredményt csak
# az admin felület mutatja meg – a dolgozó ebből semmit nem lát.

FRISSITES = {"allapot": "nincs", "info": None}


def start_update_check() -> None:
    """Induláskor, háttérszálon megnézi, van-e újabb verzió."""
    if frissito is None or not frissito.is_enabled():
        return
    FRISSITES["allapot"] = "keres"

    def kesz(info):
        FRISSITES["info"] = info
        FRISSITES["allapot"] = "van" if info else "naprakesz"

    frissito.check_async(kesz)


def resolve_workbook(parent=None) -> Path | None:
    """Megkeresi vagy bekéri az Excel-fájl elérési útját."""
    cfg = load_config()
    saved = cfg.get("workbook_path")
    if saved:
        p = Path(saved)
        if not p.is_absolute():
            p = DATA_DIR / p
        if p.exists():
            return p

    template = RESOURCE_DIR / "keszlet.xlsx"
    target = DATA_DIR / "keszlet.xlsx"
    if not target.exists() and template.exists() and template != target:
        try:
            shutil.copyfile(template, target)
        except OSError:
            pass

    for cand in (DATA_DIR / "keszlet.xlsx",
                 Path.home() / "Downloads" / "keszlet.xlsx"):
        if cand.exists():
            update_config(workbook_path=str(cand))
            return cand

    messagebox.showinfo(
        "Excel kiválasztása",
        "Nem találom a keszlet.xlsx fájlt.\n"
        "Kérlek válaszd ki a készlet Excel-fájlt.",
        parent=parent,
    )
    path = filedialog.askopenfilename(
        title="Válaszd ki a készlet Excel-fájlt",
        filetypes=[("Excel fájl", "*.xlsx"), ("Minden fájl", "*.*")],
        parent=parent,
    )
    if not path:
        return None
    update_config(workbook_path=path)
    return Path(path)

ASSETS_DIR = RESOURCE_DIR / "assets"

WINDOW_BG       = "#F1F4F1"
SURFACE         = "#FFFFFF"
HEADER_BG       = "#16352B"
HEADER_FG       = "#F4F1E8"
ACCENT          = "#2F855A"
ACCENT_HOVER    = "#276749"
SUCCESS         = "#1F7A4D"
SUCCESS_HOVER   = "#16623B"
SUCCESS_GLOW    = "#2E9C63"
SUCCESS_TINT    = "#DCF5E6"
SUCCESS_TEXT    = "#166538"
DANGER          = "#D92D20"
DANGER_HOVER    = "#B42318"
DANGER_TINT     = "#FEE4E2"
DANGER_TEXT     = "#912018"
TEXT            = "#16241E"
MUTED           = "#5B6B62"
BORDER          = "#D5DED7"
TABLE_HEADER_BG = "#1E3A30"
ROW_ALT         = "#F3F7F4"
SELECTION       = "#CDEBD9"
FOOTER_BG       = "#E4EAE5"
TOTAL_CAPTION   = "#BFE3CD"
SHADOW          = "#D5DDD7"
NEUTRAL_BTN       = "#FFFFFF"
NEUTRAL_BTN_HOVER = "#EAF0EB"

WITHDRAW          = "#B45309"
WITHDRAW_HOVER    = "#92400E"
WITHDRAW_GLOW     = "#D97706"
WITHDRAW_TINT     = "#FEF3C7"
WITHDRAW_TEXT     = "#92400E"
WITHDRAW_STRIP    = "#7C2D12"
WITHDRAW_CAPTION  = "#FCE4C4"

DEV_BAR_BG     = "#312E81"
DEV_BAR_FG     = "#E0E7FF"
DEV_BTN        = "#4338CA"
DEV_BTN_HOVER  = "#3730A3"
DEV_TEST_PATH  = "keszlet_teszt.xlsx"

def _round_rect_points(x1, y1, x2, y2, r):
    """Lekerekített téglalap pontjai create_polygon(smooth=True)-hoz."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]

class RoundButton(tk.Canvas):
    """Lekerekített, hoverre színt váltó gomb (tk.Canvas-ra rajzolva).

    A canvas háttere a szülő hátterét kapja, így a lekerekített sarkok
    tisztán illeszkednek (nincs négyzetes „glória”).
    """

    def __init__(self, master, text, *, fill, hover, fg, parent_bg,
                 width=180, height=54, radius=14,
                 font=None, command=None):
        if font is None:
            font = (FONT_UI, 14, "bold")
        super().__init__(master, width=width, height=height,
                         highlightthickness=0, bd=0, bg=parent_bg)
        self._fill, self._hover, self._command = fill, hover, command
        self._hovering = False
        self._shape = self.create_polygon(
            _round_rect_points(2, 2, width - 2, height - 2, radius),
            smooth=True, fill=fill, outline=fill)
        self._text_id = self.create_text(width // 2, height // 2, text=text,
                                         fill=fg, font=font)
        self.configure(cursor="hand2")
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonRelease-1>", self._on_click)

    def _paint(self, color):
        self.itemconfig(self._shape, fill=color, outline=color)

    def _on_enter(self, _event):
        self._hovering = True
        self._paint(self._hover)

    def _on_leave(self, _event):
        self._hovering = False
        self._paint(self._fill)

    def glow(self, color):
        """Pulzáló szín – csak ha épp nincs az egér a gombon (a hover nyer)."""
        if not self._hovering:
            self._paint(color)

    def set_text(self, text):
        self.itemconfig(self._text_id, text=text)

    def set_colors(self, fill, hover, fg=None):
        """A gomb alap-/hover-színének (és opcionálisan a felirat színének)
        cseréje futás közben – pl. mód- vagy állapotváltáskor."""
        self._fill, self._hover = fill, hover
        if fg is not None:
            self.itemconfig(self._text_id, fill=fg)
        if not self._hovering:
            self._paint(fill)

    def _on_click(self, _event):
        if self._command:
            self._command()

class Panel(tk.Frame):
    """Frame, melynek háttere egy lekerekített téglalap (canvasra rajzolva).

    A tartalom (címkék, mezők) a Frame gyermekeiként kerülnek rá, fölé.
    Átméretezésnél a háttér automatikusan újrarajzolódik.
    """

    def __init__(self, master, *, fill, outline=None, radius=16,
                 parent_bg, height=None, shadow=None):
        super().__init__(master, bg=parent_bg, height=height)
        if height:
            self.pack_propagate(False)
        self._fill = fill
        self._outline = outline or fill
        self._radius = radius
        self._shadow = shadow
        self.bg_canvas = tk.Canvas(self, highlightthickness=0, bd=0,
                                   bg=parent_bg)
        self.bg_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.bg_canvas.bind("<Configure>", self._redraw)

    def _redraw(self, _event=None):
        c = self.bg_canvas
        c.delete("bg")
        w, h = c.winfo_width(), c.winfo_height()
        if w <= 2 or h <= 2:
            return
        off = 4 if self._shadow else 1
        if self._shadow:
            c.create_polygon(
                _round_rect_points(3, 5, w - 1, h - 1, self._radius),
                smooth=True, fill=self._shadow, outline=self._shadow,
                tags="bg")
        c.create_polygon(
            _round_rect_points(1, 1, w - off, h - off, self._radius),
            smooth=True, fill=self._fill, outline=self._outline, tags="bg")
        c.tag_lower("bg")

    def set_outline(self, color):
        self._outline = color
        self._redraw()

    def set_fill(self, color):
        self._fill = color
        self._redraw()

class App(tk.Tk):
    def __init__(self, inventory: Inventory, dev: bool = False,
                 operator: str | None = None):
        super().__init__()
        self.inv = inventory
        self.operator = operator or None
        self.relogin = False
        self.cart: dict[str, list] = {}
        self.mode = "sale"

        self.dev = bool(dev)
        self.dev_dry_run = tk.BooleanVar(value=True)
        self._last_commit: dict | None = None

        self.title("Oxigén készletező")
        self.configure(bg=WINDOW_BG)
        self.geometry("1000x800")
        self.minsize(900, 660)

        self._images: dict[str, tk.PhotoImage] = {}
        self._load_assets()

        self._shown_total = 0
        self._total_anim = None
        self._pulse_phase = 0

        self._init_style()
        self._build_ui()
        self._restore_session()
        self._refresh_cart()
        self.bind("<Escape>", lambda e: self._focus_entry())
        self.bind_all("<Control-Shift-D>", lambda e: self._toggle_dev())
        self.bind_all("<Control-Shift-d>", lambda e: self._toggle_dev())
        self._apply_dev_state()
        self.after(200, self._focus_entry)
        self._fade_in()
        self._pulse_sell()

    def _load_assets(self):
        """A logó/ikon PNG-k betöltése; ha bármi hiányzik, csendben kihagyjuk."""
        files = {
            "logo": "logo_light.png",
            "icon": "app_icon.png",
            "watermark": "logo_watermark.png",
        }
        for key, fname in files.items():
            path = ASSETS_DIR / fname
            if not path.exists():
                continue
            try:
                self._images[key] = tk.PhotoImage(file=str(path))
            except tk.TclError:
                pass
        if "icon" in self._images:
            try:
                self.iconphoto(True, self._images["icon"])
            except tk.TclError:
                pass

    def _fade_in(self, step: int = 0):
        """Az ablak lágyan beúszik (ahol a platform támogatja az átlátszóságot)."""
        try:
            alpha = min(1.0, 0.0 + step * 0.12)
            self.attributes("-alpha", alpha)
            if alpha < 1.0:
                self.after(16, lambda: self._fade_in(step + 1))
        except tk.TclError:
            try:
                self.attributes("-alpha", 1.0)
            except tk.TclError:
                pass

    def _pulse_sell(self):
        """A rögzítő gomb halványan lélegzik, ha van mit rögzíteni (mód-függő szín)."""
        import math
        base = SUCCESS if self.mode == "sale" else WITHDRAW
        glow = SUCCESS_GLOW if self.mode == "sale" else WITHDRAW_GLOW
        if getattr(self, "sell_btn", None) is not None:
            if self.cart:
                self._pulse_phase += 1
                t = (math.sin(self._pulse_phase * 0.20) + 1) / 2
                self.sell_btn.glow(_blend(base, glow, t))
            else:
                self.sell_btn.glow(base)
        self.after(55, self._pulse_sell)

    def _accent(self) -> str:
        """Az aktuális mód kiemelő színe (kék-zöld eladásnál, amber elvitelnél)."""
        return ACCENT if self.mode == "sale" else WITHDRAW

    def _init_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Big.Treeview",
                        background=SURFACE, fieldbackground=SURFACE,
                        foreground=TEXT, rowheight=40,
                        font=(FONT_UI, 13), borderwidth=0)
        style.configure("Big.Treeview.Heading",
                        background=TABLE_HEADER_BG, foreground="white",
                        font=(FONT_UI, 12, "bold"),
                        padding=(10, 10), relief="flat")
        style.map("Big.Treeview.Heading",
                  background=[("active", "#334155")])
        style.map("Big.Treeview",
                  background=[("selected", SELECTION)],
                  foreground=[("selected", TEXT)])

    def _build_ui(self):
        header = tk.Frame(self, bg=HEADER_BG, height=86)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        self.header = header
        brand = tk.Frame(header, bg=HEADER_BG)
        brand.pack(side="left", padx=24, pady=12)
        if "logo" in self._images:
            tk.Label(brand, image=self._images["logo"], bg=HEADER_BG).pack(
                side="left")
            tk.Label(brand, text="Készletező", bg=HEADER_BG, fg=HEADER_FG,
                     font=(FONT_UI, 20, "bold")).pack(side="left",
                                                         padx=(18, 0))
        else:
            tk.Label(brand, text="OXIGÉN KÉSZLETEZŐ", bg=HEADER_BG, fg=HEADER_FG,
                     font=(FONT_UI, 22, "bold")).pack(side="left")
        mfr = tk.Frame(header, bg=HEADER_BG)
        mfr.pack(side="right", padx=22)
        if self.operator:
            who = tk.Frame(mfr, bg=HEADER_BG)
            who.pack(side="left", padx=(0, 18))
            tk.Label(who, text="Eladó", bg=HEADER_BG, fg=TOTAL_CAPTION,
                     font=(FONT_UI, 9)).pack(anchor="e")
            tk.Label(who, text=f"👤 {self.operator}", bg=HEADER_BG,
                     fg=HEADER_FG, font=(FONT_UI, 13, "bold")).pack(anchor="e")
        RoundButton(mfr, "Kijelentkezés", fill="#2C4A3E", hover="#3A5C4D",
                    fg=HEADER_FG, parent_bg=HEADER_BG, width=124, height=34,
                    radius=10, font=(FONT_UI, 11, "bold"),
                    command=self._logout).pack(side="left", padx=(0, 18))
        tk.Label(mfr, text="Hónap:", bg=HEADER_BG, fg=TOTAL_CAPTION,
                 font=(FONT_UI, 12)).pack(side="left", padx=(0, 8))
        self.month_combo = ttk.Combobox(
            mfr, width=13, state="readonly", font=(FONT_UI, 12),
            values=[f"{i+1}. {HU_MONTHS[i]}" for i in range(12)])
        self.month_combo.current(datetime.now().month - 1)
        self.month_combo.pack(side="left")
        self.month_combo.bind("<<ComboboxSelected>>",
                              lambda e: self._focus_entry())

        footer = tk.Frame(self, bg=FOOTER_BG, height=32)
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)
        self.status = tk.Label(footer, text=self._status_text(), bg=FOOTER_BG,
                               fg=MUTED, font=(FONT_UI, 9), anchor="w")
        self.status.pack(side="left", padx=14)
        RoundButton(footer, "Excel kiválasztása…", fill="#CBD5E1",
                    hover="#B6C2D2", fg=TEXT, parent_bg=FOOTER_BG, width=152,
                    height=24, radius=8, font=(FONT_UI, 9),
                    command=self._change_workbook).pack(side="right", padx=6,
                                                        pady=4)
        RoundButton(footer, "Készlet újratöltése", fill="#CBD5E1",
                    hover="#B6C2D2", fg=TEXT, parent_bg=FOOTER_BG, width=152,
                    height=24, radius=8, font=(FONT_UI, 9),
                    command=self._reload).pack(side="right", padx=6, pady=4)

        body = tk.Frame(self, bg=WINDOW_BG)
        body.pack(side="top", fill="both", expand=True, padx=18, pady=14)

        modebar = tk.Frame(body, bg=WINDOW_BG)
        modebar.pack(fill="x", pady=(0, 10))
        self.btn_mode_sale = RoundButton(
            modebar, "🛒  Eladás", fill=ACCENT, hover=ACCENT_HOVER, fg="white",
            parent_bg=WINDOW_BG, width=210, height=46, radius=12,
            font=(FONT_UI, 13, "bold"),
            command=lambda: self._set_mode("sale"))
        self.btn_mode_sale.pack(side="left")
        self.btn_mode_withdraw = RoundButton(
            modebar, "📦  Elvitel", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
            fg=TEXT, parent_bg=WINDOW_BG, width=210, height=46, radius=12,
            font=(FONT_UI, 13, "bold"),
            command=lambda: self._set_mode("withdraw"))
        self.btn_mode_withdraw.pack(side="left", padx=(10, 0))
        tk.Label(modebar,
                 text="Eladás = vásárlónak.   Elvitel = selejt / visszáru / "
                      "saját (eladáson kívüli készletcsökkenés).",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 9)).pack(
            side="left", padx=14)

        self.scan_panel = Panel(body, fill=SURFACE, outline=BORDER, radius=16,
                                parent_bg=WINDOW_BG, height=128, shadow=SHADOW)
        self.scan_panel.pack(fill="x")
        tk.Label(self.scan_panel, text="VONALKÓD BEOLVASÁSA", bg=SURFACE,
                 fg=MUTED, font=(FONT_UI, 11, "bold")).pack(
            anchor="w", padx=24, pady=(16, 0))
        self.entry = tk.Entry(self.scan_panel, font=(FONT_MONO, 26, "bold"),
                              relief="flat", bd=0, bg=SURFACE, fg=TEXT,
                              insertbackground=ACCENT)
        self.entry.pack(fill="x", padx=24, pady=(2, 2))
        self.entry.bind("<Return>", self._on_enter)
        self.entry.bind("<FocusIn>",
                        lambda e: self.scan_panel.set_outline(self._accent()))
        self.entry.bind("<FocusOut>",
                        lambda e: self.scan_panel.set_outline(BORDER))
        self.scan_hint = tk.Label(
            self.scan_panel,
            text="Olvasd be a vonalkódot (Enter). Üres mezőben Enter = "
                 "eladás rögzítése.",
            bg=SURFACE, fg=MUTED, font=(FONT_UI, 9))
        self.scan_hint.pack(anchor="w", padx=24, pady=(0, 12))

        self.feedback = Panel(body, fill=WINDOW_BG, outline=WINDOW_BG,
                              radius=12, parent_bg=WINDOW_BG, height=48)
        self.feedback.pack(fill="x", pady=(12, 4))
        self.fb_label = tk.Label(self.feedback, text="", bg=WINDOW_BG,
                                 fg=SUCCESS_TEXT, font=(FONT_UI, 16, "bold"),
                                 anchor="w")
        self.fb_label.pack(side="left", padx=18, pady=8)

        table_card = tk.Frame(body, bg=SURFACE, highlightbackground=BORDER,
                              highlightcolor=BORDER, highlightthickness=1, bd=0)
        table_card.pack(fill="both", expand=True)
        columns = ("name", "qty", "price", "sum", "stock")
        self.tree = ttk.Treeview(table_card, columns=columns, show="headings",
                                 selectmode="browse", style="Big.Treeview",
                                 height=5)
        self.tree.heading("name", text="  Megnevezés")
        self.tree.heading("qty", text="Db")
        self.tree.heading("price", text="Egységár")
        self.tree.heading("sum", text="Összesen")
        self.tree.heading("stock", text="Készlet utána")
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("qty", width=70, anchor="center")
        self.tree.column("price", width=150, anchor="e")
        self.tree.column("sum", width=160, anchor="e")
        self.tree.column("stock", width=140, anchor="center")
        self.tree.tag_configure("odd", background=SURFACE)
        self.tree.tag_configure("even", background=ROW_ALT)
        self.tree.tag_configure("low", foreground=DANGER,
                                font=(FONT_UI, 13, "bold"))
        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(table_card, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._remove_one_selected)
        self.tree.bind("<Delete>", self._remove_line_selected)

        self.watermark = None
        if "watermark" in self._images:
            self.watermark = tk.Label(table_card, image=self._images["watermark"],
                                      bg=SURFACE)
            self.watermark.place(relx=0.5, rely=0.5, anchor="center")

        self.total_strip = Panel(body, fill=HEADER_BG, outline=HEADER_BG,
                                 radius=18, parent_bg=WINDOW_BG, height=90,
                                 shadow=SHADOW)
        self.total_strip.pack(fill="x", pady=(12, 0))
        self.total_inner = tk.Frame(self.total_strip, bg=HEADER_BG)
        self.total_inner.pack(fill="both", expand=True, padx=28, pady=12)
        self.total_left = tk.Frame(self.total_inner, bg=HEADER_BG)
        self.total_left.pack(side="left", anchor="w")
        self.total_caption = tk.Label(self.total_left, text="ÖSSZESEN",
                                      bg=HEADER_BG, fg=TOTAL_CAPTION,
                                      font=(FONT_UI, 14, "bold"))
        self.total_caption.pack(anchor="w")
        self.count_label = tk.Label(self.total_left, text="0 tétel",
                                    bg=HEADER_BG, fg=TOTAL_CAPTION,
                                    font=(FONT_UI, 11))
        self.count_label.pack(anchor="w")
        self.total_label = tk.Label(self.total_inner, text="0 Ft",
                                    bg=HEADER_BG, fg="white",
                                    font=(FONT_UI, 40, "bold"))
        self.total_label.pack(side="right", anchor="e")

        btns = tk.Frame(body, bg=WINDOW_BG)
        btns.pack(fill="x", pady=(14, 0))
        self.sell_btn = RoundButton(
            btns, "✓  Eladás rögzítése  (Enter)", fill=SUCCESS,
            hover=SUCCESS_HOVER, fg="white", parent_bg=WINDOW_BG,
            width=346, height=56, radius=15, font=(FONT_UI, 15, "bold"),
            command=self._finalize)
        self.sell_btn.pack(side="left")
        RoundButton(btns, "−  1 db", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=116, height=52,
                    command=self._remove_one_selected).pack(side="left",
                                                            padx=(10, 0))
        RoundButton(btns, "🗑  Tétel törlése", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=WINDOW_BG,
                    width=182, height=52,
                    command=self._remove_line_selected).pack(side="left",
                                                             padx=(10, 0))
        RoundButton(btns, "Ürítés", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=WINDOW_BG,
                    width=120, height=52,
                    command=self._clear_cart).pack(side="left", padx=(10, 0))

        self._build_dev_bar()

    def _build_dev_bar(self):
        """Az indigó fejlesztői eszköztár felépítése (a láthatóságát a
        _apply_dev_state kezeli)."""
        bar = tk.Frame(self, bg=DEV_BAR_BG, height=44)
        bar.pack_propagate(False)
        self.dev_bar = bar

        tk.Label(bar, text="🧪 FEJLESZTŐI MÓD", bg=DEV_BAR_BG, fg=DEV_BAR_FG,
                 font=(FONT_UI, 11, "bold")).pack(side="left", padx=(14, 12))

        tk.Checkbutton(
            bar, text="Szárazfutás (nem ír Excelbe)", variable=self.dev_dry_run,
            bg=DEV_BAR_BG, fg=DEV_BAR_FG, selectcolor=DEV_BTN,
            activebackground=DEV_BAR_BG, activeforeground=DEV_BAR_FG,
            font=(FONT_UI, 10), bd=0, highlightthickness=0,
            command=self._update_dev_bar).pack(side="left", padx=(0, 12))

        def devbtn(text, cmd):
            return tk.Button(
                bar, text=text, command=cmd, bg=DEV_BTN, fg="white",
                activebackground=DEV_BTN_HOVER, activeforeground="white",
                font=(FONT_UI, 10, "bold"), bd=0, relief="flat",
                padx=10, pady=2, cursor="hand2")

        devbtn("🎲 Random", self._dev_random_scan).pack(side="left", padx=3, pady=7)
        devbtn("📋 Lista…", self._dev_pick_product).pack(side="left", padx=3, pady=7)
        devbtn("↩ Visszavonás", self._dev_undo_last).pack(side="left", padx=3, pady=7)
        devbtn("🧪 Teszt-fájl", self._dev_use_test_file).pack(side="left", padx=3, pady=7)
        devbtn("🔎 Diagnosztika", self._dev_show_diag).pack(side="left", padx=3, pady=7)

        self.dev_info = tk.Label(bar, text="", bg=DEV_BAR_BG, fg=DEV_BAR_FG,
                                 font=(FONT_MONO, 9))
        self.dev_info.pack(side="right", padx=14)

    def _apply_dev_state(self):
        """A dev-sáv láthatóságának és az ablak címének összehangolása."""
        if self.dev:
            self.dev_bar.pack(fill="x", after=self.header)
            self.title("Oxigén készletező  —  🧪 FEJLESZTŐI MÓD")
            self._update_dev_bar()
        else:
            self.dev_bar.pack_forget()
            self.title("Oxigén készletező")

    def _toggle_dev(self):
        self.dev = not self.dev
        self._apply_dev_state()
        self._set_feedback(
            "🧪 Fejlesztői mód BE (Ctrl+Shift+D a kikapcsoláshoz)."
            if self.dev else "Fejlesztői mód KI.")
        self._focus_entry()

    def _update_dev_bar(self):
        if not getattr(self, "dev_info", None):
            return
        dry = "SZÁRAZ" if self.dev_dry_run.get() else "ÉLES ÍRÁS"
        self.dev_info.configure(
            text=f"[{dry}]  {len(self.inv.products)} db  •  "
                 f"{Path(self.inv.path).name}")

    def _dev_random_scan(self):
        import random
        pool = [p for p in self.inv.products if p.barcode]
        if not pool:
            self._set_feedback("DEV: nincs vonalkódos termék.", error=True)
            return
        self._scan(random.choice(pool).barcode)

    def _dev_pick_product(self):
        """Kereshető terméklista – a kiválasztott termék vonalkódját „beolvassa”."""
        pool = [p for p in self.inv.products if p.barcode]
        if not pool:
            self._set_feedback("DEV: nincs vonalkódos termék.", error=True)
            return
        dlg = tk.Toplevel(self)
        dlg.title("DEV – termék kiválasztása")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        tk.Label(dlg, text="Szűrés (név vagy vonalkód):", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 11)).pack(anchor="w", padx=16, pady=(14, 2))
        ent = tk.Entry(dlg, font=(FONT_UI, 13))
        ent.pack(padx=16, fill="x")
        lb = tk.Listbox(dlg, font=(FONT_MONO, 11), height=16, width=58,
                        activestyle="none")
        lb.pack(padx=16, pady=12, fill="both", expand=True)

        shown: list = []

        def refill(*_):
            q = ent.get().strip().lower()
            lb.delete(0, "end")
            shown.clear()
            for p in pool:
                if q in f"{p.label} {p.barcode}".lower():
                    shown.append(p)
                    lb.insert("end", f"{(p.barcode or ''):<16} {p.label}  "
                                     f"(készlet: {int(p.stock or 0)})")
            if shown:
                lb.selection_set(0)

        def choose(*_):
            sel = lb.curselection()
            if not sel:
                return
            p = shown[sel[0]]
            dlg.destroy()
            self._scan(p.barcode)

        ent.bind("<KeyRelease>", refill)
        ent.bind("<Return>", choose)
        lb.bind("<Double-1>", choose)
        lb.bind("<Return>", choose)
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        refill()
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + 80
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        ent.focus_set()
        dlg.grab_set()

    def _dev_undo_last(self):
        lc = self._last_commit
        if not lc:
            self._set_feedback(
                "DEV: nincs visszavonható (éles) rögzítés.", error=True)
            return
        if lc["mode"] == "sale":
            ok, _ = self._safe_commit(
                lambda: self.inv.undo_sale(lc["items"], lc["month"]))
        else:
            ok, _ = self._safe_commit(
                lambda: self.inv.undo_withdrawal(
                    lc["items"], lc["month"], lc.get("stamp")))
        if not ok:
            return
        self._last_commit = None
        self._refresh_cart()
        self.status.configure(text=self._status_text())
        self._update_dev_bar()
        self._set_feedback("↩ DEV: utolsó rögzítés visszavonva (készlet visszaállt).")

    def _dev_use_test_file(self):
        """Biztonságos teszteléshez egy másolatra vált (keszlet_teszt.xlsx),
        így az ÉLES írás sem a valódi fájlt módosítja."""
        src = Path(self.inv.path)
        if src.name == DEV_TEST_PATH:
            self._set_feedback("DEV: már a teszt-fájlon dolgozol.")
            return
        test = src.parent / DEV_TEST_PATH
        try:
            if not test.exists() or messagebox.askyesno(
                    "Teszt-fájl",
                    f"Létezik már: {test.name}\n"
                    f"Felülírjam a jelenlegi adatok másolatával?", parent=self):
                shutil.copy2(src, test)
            new_inv = Inventory(test)
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült a teszt-fájl:\n{exc}",
                                 parent=self)
            return
        self.inv = new_inv
        self.cart.clear()
        self._last_commit = None
        self._refresh_cart()
        self.status.configure(text=self._status_text())
        self._update_dev_bar()
        self._set_feedback(f"🧪 DEV: átváltva a teszt-fájlra ({test.name}). "
                           f"Az éles fájl érintetlen marad.")

    def _dev_show_diag(self):
        inv = self.inv
        cols = "\n".join(f"    {k}: {v}" for k, v in sorted(inv.cols.items()))
        samples = "\n".join(
            f"    {(p.barcode or '—'):<16}{p.label[:30]:<32}készlet={int(p.stock or 0)}"
            for p in inv.products[:8]) or "    (nincs termék)"
        text = (
            f"Fájl:         {inv.path}\n"
            f"Munkalap:     {inv.sheet_name}\n"
            f"Termékek:     {len(inv.products)}\n"
            f"Szárazfutás:  {'BE' if self.dev_dry_run.get() else 'KI'}\n\n"
            f"Blokk-kezdő oszlopok:\n"
            f"    beszerzés:  {inv.purchase_start_col}\n"
            f"    eladás:     {inv.sale_start_col}\n"
            f"    elvitel:    {inv.withdraw_start_col or '(nincs)'}\n\n"
            f"Felismert oszlopok:\n{cols}\n\n"
            f"Minta termékek (max 8):\n{samples}")
        messagebox.showinfo("DEV – diagnosztika", text, parent=self)

    def _status_text(self) -> str:
        return (f"Adatbázis: {self.inv.path}    •    "
                f"{len(self.inv.products)} termék betöltve    •    "
                f"A készlet szerkesztéséhez nyisd meg az Excelt "
                f"(rögzítéshez legyen bezárva).")

    def _focus_entry(self):
        self.entry.focus_set()
        self.entry.selection_range(0, "end")

    def _selected_month(self) -> int:
        return self.month_combo.current() + 1

    def _cart_total(self) -> int:
        return sum(int(p.gross_sale or 0) * qty for p, qty in self.cart.values())

    def _on_enter(self, event=None):
        code = self.entry.get().strip()
        self.entry.delete(0, "end")
        if not code:
            self._finalize()
            return
        self._scan(code)

    def _scan(self, code: str):
        product = self.inv.find(code)
        if product is None:
            self.bell()
            self._flash_scan(False)
            self._set_feedback(
                f"✕   Ismeretlen vonalkód: {code}  –  nincs az adatbázisban!",
                error=True)
            self._offer_new_product(code)
            return
        key = product.barcode
        entry = self.cart.get(key)
        if entry is None:
            self.cart[key] = [product, 1]
        else:
            entry[1] += 1
        qty = self.cart[key][1]
        prefix = "✓" if self.mode == "sale" else "📦"
        stock = int(product.stock or 0)
        remaining = stock - qty
        base = (f"{prefix}   {product.label}    {format_ft(product.gross_sale)}"
                + (f"    ×{qty}" if qty > 1 else "")
                + f"     •  készlet: {stock} db")
        if remaining < 0:
            self._flash_scan(False)
            self._set_feedback(base + "   ⚠ nincs ennyi raktáron!", error=True)
        else:
            self._flash_scan(True)
            self._set_feedback(base)
        self._refresh_cart()

    def _set_feedback(self, text: str, error: bool = False):
        """A visszajelző pill színének és szövegének beállítása."""
        if not text:
            self.feedback.set_fill(WINDOW_BG)
            self.feedback.set_outline(WINDOW_BG)
            self.fb_label.configure(text="", bg=WINDOW_BG)
            return
        tint = DANGER_TINT if error else SUCCESS_TINT
        fg = DANGER_TEXT if error else SUCCESS_TEXT
        self.feedback.set_fill(tint)
        self.feedback.set_outline(tint)
        self.fb_label.configure(text=text, bg=tint, fg=fg)
        self._pop_feedback()

    def _pop_feedback(self):
        """Rövid 'pop': a felirat egy pillanatra megnő, majd visszaáll."""
        self.fb_label.configure(font=(FONT_UI, 19, "bold"))
        self.after(90, lambda: self.fb_label.configure(
            font=(FONT_UI, 16, "bold")))

    def _flash_scan(self, ok: bool):
        """A beolvasó mező keretét a mód színére / pirosra villantja, majd visszaáll."""
        self.scan_panel.set_outline(self._accent() if ok else DANGER)
        self.after(550, self._restore_scan_outline)

    def _restore_scan_outline(self):
        focused = self.focus_get() is self.entry
        self.scan_panel.set_outline(self._accent() if focused else BORDER)

    def _offer_new_product(self, code: str):
        """Ismeretlen vonalkódnál felajánlja a termék felvételét, és siker
        esetén rögtön a kosárba teszi (folytatódik az eladás/elvitel)."""
        data = self._ask_new_product(code)
        if not data:
            self._focus_entry()
            return
        ok, product = self._safe_commit(lambda: self.inv.add_product(**data))
        if not ok:
            return
        if product is None:
            self._set_feedback("Nem sikerült felvenni a terméket.", error=True)
            self._focus_entry()
            return
        self.cart[product.barcode] = [product, 1]
        self._flash_scan(True)
        self._set_feedback(
            f"➕   Új termék felvéve: {product.label}    "
            f"{format_ft(product.gross_sale)}   –   Excelbe mentve.")
        self.status.configure(text=self._status_text())
        self._refresh_cart()

    def _ask_new_product(self, code: str) -> dict | None:
        """Kis űrlap-ablak az új termék adataihoz. Visszaad: dict vagy None."""
        dlg = tk.Toplevel(self)
        dlg.title("Új termék felvétele")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)
        result = {"data": None}

        tk.Label(dlg, text="Ismeretlen vonalkód – új termék", bg=WINDOW_BG,
                 fg=TEXT, font=(FONT_UI, 15, "bold")).pack(padx=30, pady=(22, 2))
        tk.Label(dlg, text=f"Vonalkód: {code}", bg=WINDOW_BG, fg=ACCENT,
                 font=(FONT_MONO, 13, "bold")).pack(padx=30, pady=(0, 2))
        tk.Label(dlg, text="A termék az Excelbe is bekerül (új sorként).",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10)).pack(
            padx=30, pady=(0, 14))

        form = tk.Frame(dlg, bg=WINDOW_BG)
        form.pack(padx=30, fill="x")

        def field(label):
            tk.Label(form, text=label, bg=WINDOW_BG, fg=TEXT,
                     font=(FONT_UI, 11, "bold")).pack(anchor="w", pady=(8, 2))
            e = tk.Entry(form, font=(FONT_UI, 14), relief="flat",
                         bg=SURFACE, fg=TEXT, insertbackground=ACCENT,
                         highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=ACCENT)
            e.pack(fill="x", ipady=6)
            return e

        e_name = field("Megnevezés *")
        e_other = field("Méret / egyéb")
        e_price = field("Bruttó eladási ár (Ft) *")
        e_open = field("Készlet (db)")
        e_open.insert(0, "1")

        err = tk.Label(dlg, text="", bg=WINDOW_BG, fg=DANGER_TEXT,
                       font=(FONT_UI, 10, "bold"))
        err.pack(padx=30, pady=(8, 0))

        def submit():
            name = e_name.get().strip()
            price = parse_price(e_price.get())
            if not name:
                err.configure(text="A megnevezés kötelező.")
                e_name.focus_set()
                return
            if price is None:
                err.configure(text="Az árat számként add meg (pl. 2990).")
                e_price.focus_set()
                return
            result["data"] = {
                "name": name,
                "other": e_other.get().strip(),
                "barcode": code,
                "gross_sale": price,
                "purchase": parse_int(e_open.get()),
                "month": self._selected_month(),
            }
            dlg.destroy()

        btns = tk.Frame(dlg, bg=WINDOW_BG)
        btns.pack(padx=30, pady=(16, 22), fill="x")
        RoundButton(btns, "✓  Felvétel", fill=SUCCESS, hover=SUCCESS_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=200, height=46,
                    radius=12, font=(FONT_UI, 13, "bold"),
                    command=submit).pack(side="left")
        RoundButton(btns, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=120, height=46,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(side="right")

        e_name.bind("<Return>", lambda e: e_price.focus_set())
        e_price.bind("<Return>", lambda e: submit())
        e_open.bind("<Return>", lambda e: submit())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        e_name.focus_set()
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.grab_set()
        self.wait_window(dlg)
        return result["data"]

    def _refresh_cart(self):
        self.tree.delete(*self.tree.get_children())
        for i, (code, (p, qty)) in enumerate(self.cart.items()):
            after = int((p.stock or 0) - qty)
            zebra = "even" if i % 2 else "odd"
            tags = (zebra, "low") if after < 0 else (zebra,)
            self.tree.insert(
                "", "end", iid=code, tags=tags,
                values=(f"  {p.label}", qty, format_ft(p.gross_sale),
                        format_ft(int(p.gross_sale or 0) * qty),
                        after))
        self._animate_total(self._cart_total())
        n = sum(qty for _, qty in self.cart.values())
        self.count_label.configure(text=f"{n} tétel")
        if self.watermark is not None:
            if self.cart:
                self.watermark.place_forget()
            else:
                self.watermark.place(relx=0.5, rely=0.5, anchor="center")
        self._save_session()
        self._focus_entry()

    def _save_session(self):
        """Az aktuális (még le nem zárt) kosarat lemezre menti a mai dátummal.
        Minden kosár-változáskor lefut; rögzítés/ürítés után üres listát ment."""
        data = {
            "date": datetime.now().strftime("%Y-%m-%d"),
            "mode": self.mode,
            "workbook": str(self.inv.path),
            "items": [[code, qty] for code, (_, qty) in self.cart.items()],
        }
        try:
            SESSION_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _restore_session(self):
        """Ha van MAI, még le nem zárt munkamenet ugyanahhoz az Excelhez,
        visszatölti a kosarat (és a módot). Régebbi napit figyelmen kívül hagy."""
        if not SESSION_FILE.exists():
            return
        try:
            data = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return
        if data.get("date") != datetime.now().strftime("%Y-%m-%d"):
            return
        if data.get("workbook") and data["workbook"] != str(self.inv.path):
            return
        mode = data.get("mode", "sale")
        if mode in ("sale", "withdraw"):
            self.mode = mode
            self._apply_mode_visuals()
        restored = 0
        for entry in data.get("items", []):
            try:
                code, qty = entry[0], int(entry[1])
            except (TypeError, ValueError, IndexError):
                continue
            product = self.inv.find(code)
            if product is None or qty <= 0:
                continue
            self.cart[product.barcode] = [product, qty]
            restored += 1
        if restored:
            n = sum(qty for _, qty in self.cart.values())
            self._set_feedback(
                f"↩  Mai, le nem zárt lista visszatöltve – {n} db "
                f"({restored} tétel). Folytathatod vagy rögzítheted.")

    def _animate_total(self, target: int, frames: int = 12):
        """Az ÖSSZESEN érték lágyan 'felpörög' az új összegre."""
        if self._total_anim is not None:
            self.after_cancel(self._total_anim)
            self._total_anim = None
        start = self._shown_total

        def step(i: int):
            t = i / frames
            t = 1 - (1 - t) * (1 - t)
            val = int(round(start + (target - start) * t))
            self.total_label.configure(text=format_ft(val))
            if i < frames:
                self._total_anim = self.after(16, lambda: step(i + 1))
            else:
                self._shown_total = target
                self.total_label.configure(text=format_ft(target))
                self._total_anim = None

        if start == target:
            self.total_label.configure(text=format_ft(target))
        else:
            step(0)

    def _selected_code(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _remove_one_selected(self, event=None):
        code = self._selected_code()
        if not code or code not in self.cart:
            return
        self.cart[code][1] -= 1
        if self.cart[code][1] <= 0:
            del self.cart[code]
        self._refresh_cart()

    def _remove_line_selected(self, event=None):
        code = self._selected_code()
        if code and code in self.cart:
            del self.cart[code]
            self._refresh_cart()

    def _clear_cart(self):
        if self.cart and not messagebox.askyesno(
                "Ürítés", "Biztosan törlöd az aktuális listát?", parent=self):
            return
        self.cart.clear()
        self._set_feedback("")
        self._refresh_cart()

    def _set_mode(self, mode: str):
        if mode == self.mode:
            self._focus_entry()
            return
        if self.cart and not messagebox.askyesno(
                "Mód váltása",
                "A listában lévő tételek elvesznek a váltáskor. Folytatod?",
                parent=self):
            return
        self.mode = mode
        self._apply_mode_visuals()

        self.cart.clear()
        self._set_feedback("")
        self._refresh_cart()

    def _apply_mode_visuals(self):
        """A felület átszínezése az aktuális self.mode szerint (kosár-érintés
        nélkül – így a napi munkamenet visszatöltése is használhatja)."""
        sale = (self.mode == "sale")
        strip = HEADER_BG if sale else WITHDRAW_STRIP
        cap = TOTAL_CAPTION if sale else WITHDRAW_CAPTION

        if sale:
            self.sell_btn.set_text("✓  Eladás rögzítése  (Enter)")
            self.sell_btn.set_colors(SUCCESS, SUCCESS_HOVER, fg="white")
        else:
            self.sell_btn.set_text("📦  Elvitel rögzítése  (Enter)")
            self.sell_btn.set_colors(WITHDRAW, WITHDRAW_HOVER, fg="white")

        self.total_strip.set_fill(strip)
        self.total_strip.set_outline(strip)
        for w in (self.total_inner, self.total_left):
            w.configure(bg=strip)
        self.total_caption.configure(
            bg=strip, fg=cap, text=("ÖSSZESEN" if sale else "ELVITEL ÉRTÉKE"))
        self.count_label.configure(bg=strip, fg=cap)
        self.total_label.configure(bg=strip)

        if sale:
            self.btn_mode_sale.set_colors(ACCENT, ACCENT_HOVER, fg="white")
            self.btn_mode_withdraw.set_colors(NEUTRAL_BTN, NEUTRAL_BTN_HOVER,
                                              fg=TEXT)
        else:
            self.btn_mode_sale.set_colors(NEUTRAL_BTN, NEUTRAL_BTN_HOVER, fg=TEXT)
            self.btn_mode_withdraw.set_colors(WITHDRAW, WITHDRAW_HOVER, fg="white")

        self.scan_hint.configure(text=(
            "Olvasd be a vonalkódot (Enter). Üres mezőben Enter = "
            "eladás rögzítése." if sale else
            "ELVITEL mód – olvasd be a kivett termékeket. Üres mezőben "
            "Enter = elvitel rögzítése."))
        self._restore_scan_outline()

    def _finalize(self):
        """Az aktuális mód rögzítése: eladás vagy elvitel."""
        if self.mode == "withdraw":
            self._finalize_withdraw()
        else:
            self._finalize_sale()

    def _safe_commit(self, fn):
        """A megadott rögzítő-hívás futtatása barátságos hibakezeléssel.

        Visszaad: (siker?, hiányzó_vonalkódok). Hibánál (False, []).
        """
        try:
            return True, fn()
        except PermissionError:
            messagebox.showerror(
                "Az Excel meg van nyitva",
                "Nem tudok írni a fájlba, mert valószínűleg meg van nyitva "
                "Excelben.\n\nZárd be az Excelt, majd próbáld újra. "
                "A lista megmarad.",
                parent=self)
        except FileNotFoundError:
            messagebox.showerror(
                "Hiányzó fájl",
                f"Nem található az Excel-fájl:\n{self.inv.path}",
                parent=self)
        except Exception as exc:
            messagebox.showerror(
                "Hiba a rögzítéskor", f"Váratlan hiba:\n{exc}", parent=self)
        self._focus_entry()
        return False, []

    def _warn_missing(self, missing):
        if missing:
            messagebox.showwarning(
                "Figyelem",
                "Néhány terméket nem találtam az Excelben (talán törölve "
                "lett):\n" + ", ".join(missing),
                parent=self)

    def _finalize_withdraw(self):
        if not self.cart:
            self._set_feedback(
                "A lista üres – előbb olvasd be a kivett termékeket.",
                error=True)
            return

        n_items = sum(qty for _, qty in self.cart.values())
        month = self._selected_month()
        items = [(code, qty) for code, (_, qty) in self.cart.items()]

        if self.dev and self.dev_dry_run.get():
            self._warn_missing([c for c, _ in items if self.inv.find(c) is None])
            self.cart.clear()
            self._refresh_cart()
            self._set_feedback(
                f"🧪 DEV szárazfutás: NEM írtam Excelbe – {n_items} db elvitel "
                f"levonva/naplózva LENNE.")
            self._update_dev_bar()
            return

        reason = self._ask_reason()
        if reason is None:
            self._focus_entry()
            return

        ok, missing = self._safe_commit(
            lambda: self.inv.commit_withdrawal(items, month, reason))
        if not ok:
            return
        self._warn_missing(missing)
        self._last_commit = {"mode": "withdraw", "items": items, "month": month,
                             "reason": reason, "stamp": self.inv.last_log_stamp}

        self.cart.clear()
        self._refresh_cart()
        self._set_feedback(
            f"📦  Elvitel rögzítve – {n_items} db  ({reason})  "
            f"– levonva a készletből, naplózva.")
        self.status.configure(text=self._status_text())
        self._update_dev_bar()

    def _ask_reason(self) -> str | None:
        """Kis modális ablak: az elvitel okának kiválasztása (gombokkal)."""
        dlg = tk.Toplevel(self)
        dlg.title("Elvitel oka")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)
        result = {"reason": None}

        tk.Label(dlg, text="Mi az elvitel oka?", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 15, "bold")).pack(padx=30, pady=(24, 4))
        tk.Label(dlg, text="A választott ok bekerül a Napló lapra is.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10)).pack(
            padx=30, pady=(0, 16))

        def choose(r):
            result["reason"] = r
            dlg.destroy()

        for r in WITHDRAW_REASONS:
            RoundButton(dlg, r, fill=WITHDRAW, hover=WITHDRAW_HOVER, fg="white",
                        parent_bg=WINDOW_BG, width=320, height=48, radius=12,
                        font=(FONT_UI, 13, "bold"),
                        command=lambda rr=r: choose(rr)).pack(padx=30, pady=5)
        RoundButton(dlg, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=320, height=42,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(padx=30, pady=(10, 22))
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.grab_set()
        self.wait_window(dlg)
        return result["reason"]

    def _finalize_sale(self):
        if not self.cart:
            self._set_feedback("A lista üres – előbb olvass be terméket.",
                               error=True)
            return

        total = self._cart_total()
        n_items = sum(qty for _, qty in self.cart.values())
        month = self._selected_month()
        items = [(code, qty) for code, (_, qty) in self.cart.items()]

        if self.dev and self.dev_dry_run.get():
            self._warn_missing([c for c, _ in items if self.inv.find(c) is None])
            self.cart.clear()
            self._refresh_cart()
            self._set_feedback(
                f"🧪 DEV szárazfutás: NEM írtam Excelbe – {format_ft(total)} "
                f"({n_items} db) levonva LENNE.")
            self._update_dev_bar()
            return

        lines = "\n".join(
            f"  • {p.label}  ×{qty}  =  {format_ft(int(p.gross_sale or 0) * qty)}"
            for p, qty in self.cart.values())
        if not messagebox.askyesno(
                "Eladás rögzítése",
                f"{lines}\n\n"
                f"Összesen ({n_items} db):  {format_ft(total)}\n\n"
                f"Rögzítem a(z) {HU_MONTHS[month-1]} hónaphoz, és levonom "
                f"a készletből?",
                parent=self):
            self._focus_entry()
            return

        seller = self.operator or "—"
        ok, missing = self._safe_commit(
            lambda: self.inv.commit_sale(items, month, seller=seller))
        if not ok:
            return
        self._warn_missing(missing)
        self._last_commit = {"mode": "sale", "items": items, "month": month}

        self.cart.clear()
        self._refresh_cart()
        self._set_feedback(f"✓  Eladás rögzítve – {format_ft(total)}  "
                           f"(levonva a készletből).")
        self.status.configure(text=self._status_text())
        self._update_dev_bar()

    def _reload(self):
        try:
            self.inv.load()
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült újratölteni:\n{exc}",
                                 parent=self)
            return
        self._set_feedback("Készlet újratöltve az Excelből.")
        self.status.configure(text=self._status_text())
        self._update_dev_bar()
        self._focus_entry()

    def _logout(self):
        """Kijelentkezés: visszatérés a belépő képernyőre (a kosár megmarad a
        napi munkamenet-fájlban, ahogy eddig is)."""
        if self.cart and not messagebox.askyesno(
                "Kijelentkezés",
                "Van még tétel a kosárban. Biztosan kijelentkezel?\n"
                "(A kosár megmarad, legközelebb visszatöltődik.)",
                parent=self):
            return
        self.relogin = True
        self.destroy()

    def _change_workbook(self):
        path = filedialog.askopenfilename(
            title="Válaszd ki a készlet Excel-fájlt",
            filetypes=[("Excel fájl", "*.xlsx"), ("Minden fájl", "*.*")],
            parent=self)
        if not path:
            return
        try:
            new_inv = Inventory(Path(path))
        except Exception as exc:
            messagebox.showerror("Hiba",
                                 f"Nem sikerült megnyitni a fájlt:\n{exc}",
                                 parent=self)
            return
        self.inv = new_inv
        update_config(workbook_path=path)
        self.cart.clear()
        self._last_commit = None
        self._refresh_cart()
        self.status.configure(text=self._status_text())
        self._update_dev_bar()
        self._set_feedback("Új adatbázis betöltve.")

def _load_window_icon(win) -> dict:
    """Az ablakikon és a logó betöltése (hibatűrően). Visszaadja a képeket,
    hogy a hívó eltárolhassa (különben a GC eldobná őket)."""
    images: dict[str, tk.PhotoImage] = {}
    for key, fname in (("icon", "app_icon.png"), ("logo", "logo_light.png")):
        path = ASSETS_DIR / fname
        if path.exists():
            try:
                images[key] = tk.PhotoImage(file=str(path))
            except tk.TclError:
                pass
    if "icon" in images:
        try:
            win.iconphoto(True, images["icon"])
        except tk.TclError:
            pass
    return images

def _center_window(win, w: int, h: int) -> None:
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    x = max(0, (sw - w) // 2)
    y = max(0, (sh - h) // 3)
    win.geometry(f"{w}x{h}+{x}+{y}")

def _fade_window(win, step: int = 0) -> None:
    """Lágy beúszás (ahol a platform támogatja az átlátszóságot)."""
    try:
        alpha = min(1.0, step * 0.14)
        win.attributes("-alpha", alpha)
        if alpha < 1.0:
            win.after(16, lambda: _fade_window(win, step + 1))
    except tk.TclError:
        pass

class LoginWindow(tk.Tk):
    """Indító képernyő: a felhasználó eldönti, dolgozóként vagy adminként lép be.

    Az eredmény a `result` mezőben:
        ("employee", név|None)  – dolgozói kassza
        ("admin", None)         – admin felület
        None                    – kilépés (az ablak bezárása választás nélkül)
    """

    def __init__(self):
        super().__init__()
        self.result = None
        self.title("Oxigén készletező – belépés")
        self.configure(bg=WINDOW_BG)
        self.resizable(False, False)
        self._images = _load_window_icon(self)
        self._build()
        _center_window(self, 540, 600)
        _fade_window(self)
        self.bind("<Escape>", lambda e: self.destroy())

    def _build(self):
        header = tk.Frame(self, bg=HEADER_BG, height=120)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        if "logo" in self._images:
            tk.Label(header, image=self._images["logo"], bg=HEADER_BG).pack(
                pady=(26, 0))
        else:
            tk.Label(header, text="OXIGÉN", bg=HEADER_BG, fg=HEADER_FG,
                     font=(FONT_UI, 30, "bold")).pack(pady=(30, 0))
        tk.Label(header, text="KÉSZLETEZŐ", bg=HEADER_BG, fg=TOTAL_CAPTION,
                 font=(FONT_UI, 13, "bold")).pack()

        body = tk.Frame(self, bg=WINDOW_BG)
        body.pack(fill="both", expand=True, padx=40, pady=10)

        tk.Label(body, text="Belépés mint…", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 22, "bold")).pack(pady=(26, 6))
        tk.Label(body, text="Válaszd ki, milyen módban szeretnél dolgozni.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 11)).pack(pady=(0, 26))

        RoundButton(body, "🛒   Dolgozó  (kassza)", fill=SUCCESS,
                    hover=SUCCESS_HOVER, fg="white", parent_bg=WINDOW_BG,
                    width=420, height=92, radius=18,
                    font=(FONT_UI, 18, "bold"),
                    command=self._employee).pack(pady=10)
        tk.Label(body, text="Eladás és elvitel rögzítése vonalkódolvasóval.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10)).pack(pady=(0, 18))

        RoundButton(body, "🔑   Admin  (kezelés)", fill=HEADER_BG,
                    hover="#21483B", fg="white", parent_bg=WINDOW_BG,
                    width=420, height=82, radius=18,
                    font=(FONT_UI, 17, "bold"),
                    command=self._admin).pack(pady=4)
        tk.Label(body, text="Termékek, statisztika, napló és dolgozók kezelése "
                            "(jelszóval).",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10)).pack(pady=(2, 0))

        footer = tk.Frame(self, bg=WINDOW_BG)
        footer.pack(side="bottom", fill="x", pady=(0, 14))
        RoundButton(footer, "Kilépés", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=130, height=38,
                    radius=10, font=(FONT_UI, 11),
                    command=self.destroy).pack()
        tk.Label(footer, text=f"v{APP_VERSION}", bg=WINDOW_BG, fg=MUTED,
                 font=(FONT_UI, 9)).pack(pady=(6, 0))

    def _employee(self):
        names = get_employees()
        if not names:
            self.result = ("employee", None)
            self.destroy()
            return
        name = self._pick_employee(names)
        if name is None:
            return
        self.result = ("employee", name)
        self.destroy()

    def _pick_employee(self, names: list[str]) -> str | None:
        dlg = tk.Toplevel(self)
        dlg.title("Ki dolgozik most?")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)
        chosen = {"name": None}

        tk.Label(dlg, text="Ki dolgozik most?", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 16, "bold")).pack(padx=30, pady=(22, 4))
        tk.Label(dlg, text="A neved bekerül az eladás-naplóba.", bg=WINDOW_BG,
                 fg=MUTED, font=(FONT_UI, 10)).pack(padx=30, pady=(0, 14))

        def choose(n):
            chosen["name"] = n
            dlg.destroy()

        wrap = tk.Frame(dlg, bg=WINDOW_BG)
        wrap.pack(padx=30, pady=(0, 8))
        for n in names:
            RoundButton(wrap, n, fill=SUCCESS, hover=SUCCESS_HOVER, fg="white",
                        parent_bg=WINDOW_BG, width=320, height=46, radius=12,
                        font=(FONT_UI, 13, "bold"),
                        command=lambda nn=n: choose(nn)).pack(pady=5)
        RoundButton(dlg, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=320, height=40,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(padx=30, pady=(8, 22))
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_window(dlg, 380, 180 + 56 * len(names))
        dlg.grab_set()
        self.wait_window(dlg)
        return chosen["name"]

    def _admin(self):
        pw = self._ask_password()
        if pw is None:
            return
        if not check_admin_password(pw):
            messagebox.showerror("Hibás jelszó",
                                 "A megadott admin jelszó nem helyes.",
                                 parent=self)
            return
        self.result = ("admin", None)
        self.destroy()

    def _ask_password(self) -> str | None:
        dlg = tk.Toplevel(self)
        dlg.title("Admin belépés")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)
        out = {"pw": None}

        tk.Label(dlg, text="🔑  Admin belépés", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 16, "bold")).pack(padx=30, pady=(22, 4))
        tk.Label(dlg, text="Add meg az admin jelszót.", bg=WINDOW_BG, fg=MUTED,
                 font=(FONT_UI, 10)).pack(padx=30, pady=(0, 14))
        ent = tk.Entry(dlg, show="•", font=(FONT_UI, 16), justify="center",
                       relief="flat", bg=SURFACE, fg=TEXT,
                       insertbackground=ACCENT, width=20)
        ent.pack(padx=30, ipady=6)
        ent.focus_set()

        def ok():
            out["pw"] = ent.get()
            dlg.destroy()

        bar = tk.Frame(dlg, bg=WINDOW_BG)
        bar.pack(padx=30, pady=(18, 22))
        RoundButton(bar, "Belépés", fill=HEADER_BG, hover="#21483B", fg="white",
                    parent_bg=WINDOW_BG, width=150, height=44, radius=12,
                    font=(FONT_UI, 13, "bold"), command=ok).pack(side="left",
                                                                 padx=6)
        RoundButton(bar, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=120, height=44,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(side="left", padx=6)
        ent.bind("<Return>", lambda e: ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_window(dlg, 360, 250)
        dlg.grab_set()
        self.wait_window(dlg)
        return out["pw"]

class AdminWindow(tk.Tk):
    """Admin felület füleken: Termékek, Statisztika, Napi zárás, Napló, Dolgozók."""

    def __init__(self, inventory: Inventory):
        super().__init__()
        self.inv = inventory
        self.relogin = False
        self._all_products: list[Product] = []

        self.title("Oxigén készletező – Admin")
        self.configure(bg=WINDOW_BG)
        self.geometry("1120x800")
        self.minsize(980, 660)
        self._images = _load_window_icon(self)
        self._init_style()
        self._build_ui()
        self.after(120, self._refresh_all)
        _fade_window(self)

    def _init_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Admin.Treeview", background=SURFACE,
                        fieldbackground=SURFACE, foreground=TEXT,
                        rowheight=30, font=(FONT_UI, 11), borderwidth=0)
        style.configure("Admin.Treeview.Heading", background=TABLE_HEADER_BG,
                        foreground="white", font=(FONT_UI, 11, "bold"),
                        padding=(8, 8), relief="flat")
        style.map("Admin.Treeview.Heading", background=[("active", "#334155")])
        style.map("Admin.Treeview", background=[("selected", SELECTION)],
                  foreground=[("selected", TEXT)])
        style.configure("Admin.TNotebook", background=WINDOW_BG, borderwidth=0)
        style.configure("Admin.TNotebook.Tab", font=(FONT_UI, 12, "bold"),
                        padding=(20, 10))
        style.map("Admin.TNotebook.Tab",
                  background=[("selected", SURFACE), ("!selected", "#DCE6DE")],
                  foreground=[("selected", ACCENT), ("!selected", MUTED)])

    def _build_ui(self):
        header = tk.Frame(self, bg=HEADER_BG, height=78)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        brand = tk.Frame(header, bg=HEADER_BG)
        brand.pack(side="left", padx=22, pady=10)
        if "logo" in self._images:
            tk.Label(brand, image=self._images["logo"], bg=HEADER_BG).pack(
                side="left")
        tk.Label(brand, text="Admin felület", bg=HEADER_BG, fg=HEADER_FG,
                 font=(FONT_UI, 19, "bold")).pack(side="left", padx=(16, 0))
        RoundButton(header, "Kijelentkezés", fill="#2C4A3E", hover="#3A5C4D",
                    fg=HEADER_FG, parent_bg=HEADER_BG, width=132, height=36,
                    radius=10, font=(FONT_UI, 11, "bold"),
                    command=self._logout).pack(side="right", padx=22)

        nb = ttk.Notebook(self, style="Admin.TNotebook")
        nb.pack(fill="both", expand=True, padx=14, pady=12)

        self.tab_products = tk.Frame(nb, bg=WINDOW_BG)
        self.tab_stats = tk.Frame(nb, bg=WINDOW_BG)
        self.tab_daily = tk.Frame(nb, bg=WINDOW_BG)
        self.tab_log = tk.Frame(nb, bg=WINDOW_BG)
        self.tab_staff = tk.Frame(nb, bg=WINDOW_BG)
        self.tab_update = tk.Frame(nb, bg=WINDOW_BG)
        nb.add(self.tab_products, text="  📦  Termékek  ")
        nb.add(self.tab_stats, text="  📊  Statisztika  ")
        nb.add(self.tab_daily, text="  📅  Napi zárás  ")
        nb.add(self.tab_log, text="  📒  Napló  ")
        nb.add(self.tab_staff, text="  👥  Dolgozók  ")
        nb.add(self.tab_update, text="  ⬆  Frissítés  ")
        self._nb = nb
        self._update_tab_index = 5

        self._build_products_tab()
        self._build_stats_tab()
        self._build_daily_tab()
        self._build_log_tab()
        self._build_staff_tab()
        self._build_update_tab()

    def _logout(self):
        self.relogin = True
        self.destroy()

    def _refresh_all(self):
        self._reload_inventory(silent=True)
        self._refresh_products()
        self._refresh_stats()
        self._refresh_daily()
        self._refresh_log()
        self._refresh_staff()
        self._refresh_update()

    def _reload_inventory(self, silent: bool = False):
        try:
            self.inv.load()
        except Exception as exc:
            if not silent:
                messagebox.showerror("Hiba",
                                     f"Nem sikerült újratölteni:\n{exc}",
                                     parent=self)
            return False
        return True

    def _build_products_tab(self):
        t = self.tab_products
        bar = tk.Frame(t, bg=WINDOW_BG)
        bar.pack(fill="x", padx=8, pady=(10, 6))
        tk.Label(bar, text="Keresés:", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 11)).pack(side="left")
        self.prod_search = tk.Entry(bar, font=(FONT_UI, 12), width=28,
                                    relief="flat", bg=SURFACE, fg=TEXT,
                                    insertbackground=ACCENT)
        self.prod_search.pack(side="left", padx=(8, 0), ipady=4)
        self.prod_search.bind("<KeyRelease>", lambda e: self._refresh_products())

        RoundButton(bar, "🔄  Frissítés", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=WINDOW_BG,
                    width=130, height=40, radius=10, font=(FONT_UI, 11),
                    command=self._refresh_all).pack(side="right", padx=(8, 0))
        RoundButton(bar, "🗑  Törlés", fill=DANGER, hover=DANGER_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=120, height=40,
                    radius=10, font=(FONT_UI, 12, "bold"),
                    command=self._delete_product).pack(side="right", padx=(8, 0))
        RoundButton(bar, "✏  Szerkesztés", fill=ACCENT, hover=ACCENT_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=150, height=40,
                    radius=10, font=(FONT_UI, 12, "bold"),
                    command=self._edit_product).pack(side="right", padx=(8, 0))
        RoundButton(bar, "➕  Új termék", fill=SUCCESS, hover=SUCCESS_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=140, height=40,
                    radius=10, font=(FONT_UI, 12, "bold"),
                    command=self._new_product).pack(side="right")

        card = tk.Frame(t, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1)
        card.pack(fill="both", expand=True, padx=8, pady=(4, 10))
        cols = ("name", "other", "barcode", "price", "stock")
        tv = ttk.Treeview(card, columns=cols, show="headings",
                          style="Admin.Treeview", selectmode="browse")
        for key, txt, w, anc in (
                ("name", "Megnevezés", 320, "w"),
                ("other", "Méret / egyéb", 140, "w"),
                ("barcode", "Vonalkód", 160, "w"),
                ("price", "Eladási ár", 140, "e"),
                ("stock", "Készlet", 110, "center")):
            tv.heading(key, text=txt)
            tv.column(key, width=w, anchor=anc)
        tv.tag_configure("even", background=ROW_ALT)
        tv.tag_configure("low", foreground=DANGER)
        tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(card, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tv.bind("<Double-1>", lambda e: self._edit_product())
        self.prod_tree = tv
        self._prod_by_iid: dict[str, Product] = {}

    def _refresh_products(self):
        tv = self.prod_tree
        tv.delete(*tv.get_children())
        self._prod_by_iid.clear()
        q = self.prod_search.get().strip().lower()
        for i, p in enumerate(self.inv.products):
            hay = f"{p.name} {p.other} {p.barcode}".lower()
            if q and q not in hay:
                continue
            stock = int(p.stock or 0)
            tags = ["even"] if i % 2 else []
            if stock <= 0:
                tags.append("low")
            iid = tv.insert("", "end", values=(
                p.name, p.other, p.barcode,
                format_ft(int(p.gross_sale or 0)), stock), tags=tuple(tags))
            self._prod_by_iid[iid] = p

    def _selected_product(self) -> Product | None:
        sel = self.prod_tree.selection()
        if not sel:
            return None
        return self._prod_by_iid.get(sel[0])

    def _new_product(self):
        data = self._product_dialog("Új termék felvétele", None)
        if data is None:
            return
        if self.inv.find(data["barcode"]):
            messagebox.showwarning(
                "Létező vonalkód",
                "Ezzel a vonalkóddal már van termék. Használd a Szerkesztést.",
                parent=self)
            return
        try:
            self.inv.add_product(
                name=data["name"], other=data["other"],
                barcode=data["barcode"], gross_sale=data["gross_sale"],
                purchase=data["purchase"])
        except PermissionError:
            self._excel_open_error()
            return
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült menteni:\n{exc}",
                                 parent=self)
            return
        self._refresh_products()
        self._refresh_stats()
        messagebox.showinfo("Kész", "Az új termék felkerült a készletbe.",
                            parent=self)

    def _edit_product(self):
        p = self._selected_product()
        if p is None:
            messagebox.showinfo("Nincs kijelölve",
                                "Előbb válassz ki egy terméket a listából.",
                                parent=self)
            return
        data = self._product_dialog("Termék szerkesztése", p)
        if data is None:
            return
        clash = self.inv.find(data["barcode"])
        if clash is not None and clash.row != p.row:
            messagebox.showwarning(
                "Létező vonalkód",
                "Ezt a vonalkódot már egy másik termék használja.",
                parent=self)
            return
        try:
            self.inv.update_product(
                row=p.row, name=data["name"], other=data["other"],
                barcode=data["barcode"], gross_sale=data["gross_sale"],
                opening=data["opening"])
        except PermissionError:
            self._excel_open_error()
            return
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült menteni:\n{exc}",
                                 parent=self)
            return
        self._refresh_products()
        self._refresh_stats()

    def _delete_product(self):
        p = self._selected_product()
        if p is None:
            messagebox.showinfo("Nincs kijelölve",
                                "Előbb válassz ki egy terméket a listából.",
                                parent=self)
            return
        if not messagebox.askyesno(
                "Termék törlése",
                f"Biztosan törlöd ezt a terméket?\n\n"
                f"  {p.label}\n"
                f"  Vonalkód: {p.barcode}\n"
                f"  Készlet: {int(p.stock or 0)} db\n\n"
                "A sor az Excelből is törlődik, a havi beszerzés/eladás "
                "számokkal együtt.\n(Mentés előtt biztonsági másolat készül "
                "a backup mappába.)",
                parent=self, icon="warning", default="no"):
            return
        try:
            self.inv.delete_product(row=p.row)
        except PermissionError:
            self._excel_open_error()
            return
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült törölni:\n{exc}",
                                 parent=self)
            return
        self._refresh_products()
        self._refresh_stats()

    def _excel_open_error(self):
        messagebox.showerror(
            "Az Excel meg van nyitva",
            "Nem tudok írni a fájlba, mert valószínűleg meg van nyitva "
            "Excelben.\nZárd be az Excelt, majd próbáld újra.", parent=self)

    def _product_dialog(self, title: str, prod: Product | None) -> dict | None:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)
        out = {"data": None}

        tk.Label(dlg, text=title, bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 15, "bold")).grid(row=0, column=0, columnspan=2,
                                                  padx=24, pady=(20, 14),
                                                  sticky="w")
        fields = [
            ("Megnevezés *", prod.name if prod else ""),
            ("Méret / egyéb", prod.other if prod else ""),
            ("Vonalkód *", prod.barcode if prod else ""),
            ("Bruttó eladási ár (Ft) *",
             str(int(prod.gross_sale or 0)) if prod else ""),
            ("Nyitó készlet (db)" if prod else "Készlet (db)",
             str(int(prod.opening or 0)) if prod else "0"),
        ]
        entries = []
        for i, (lab, val) in enumerate(fields, start=1):
            tk.Label(dlg, text=lab, bg=WINDOW_BG, fg=MUTED,
                     font=(FONT_UI, 11)).grid(row=i, column=0, padx=(24, 10),
                                              pady=6, sticky="e")
            e = tk.Entry(dlg, font=(FONT_UI, 13), width=26, relief="flat",
                         bg=SURFACE, fg=TEXT, insertbackground=ACCENT)
            e.insert(0, val)
            e.grid(row=i, column=1, padx=(0, 24), pady=6, ipady=4, sticky="w")
            entries.append(e)
        entries[0].focus_set()

        def save():
            name = entries[0].get().strip()
            other = entries[1].get().strip()
            barcode = norm_barcode(entries[2].get())
            price = parse_price(entries[3].get())
            qty = parse_int(entries[4].get())
            if not name or not barcode or price is None:
                messagebox.showwarning(
                    "Hiányzó adat",
                    "A Megnevezés, a Vonalkód és az Eladási ár kötelező.",
                    parent=dlg)
                return
            out["data"] = {"name": name, "other": other, "barcode": barcode,
                           "gross_sale": float(price)}
            if prod is None:
                out["data"]["purchase"] = int(qty)
            else:
                out["data"]["opening"] = int(qty)
            dlg.destroy()

        bar = tk.Frame(dlg, bg=WINDOW_BG)
        bar.grid(row=len(fields) + 1, column=0, columnspan=2, pady=(16, 20))
        RoundButton(bar, "💾  Mentés", fill=SUCCESS, hover=SUCCESS_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=160, height=44,
                    radius=12, font=(FONT_UI, 13, "bold"),
                    command=save).pack(side="left", padx=6)
        RoundButton(bar, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=120, height=44,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(side="left", padx=6)
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_window(dlg, 460, 430)
        dlg.grab_set()
        self.wait_window(dlg)
        return out["data"]

    def _build_stats_tab(self):
        t = self.tab_stats
        self.stats_cards = tk.Frame(t, bg=WINDOW_BG)
        self.stats_cards.pack(fill="x", padx=8, pady=(12, 6))

        mid = tk.Frame(t, bg=WINDOW_BG)
        mid.pack(fill="both", expand=True, padx=8, pady=(4, 10))
        mid.columnconfigure(0, weight=1, uniform="s")
        mid.columnconfigure(1, weight=1, uniform="s")
        mid.rowconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)

        self.stats_month_tree = self._mk_stats_tree(
            mid, "Havi forgalom", ("Hónap", "Eladott db", "Bevétel"),
            (140, 120, 160), row=0, col=0)
        self.stats_top_tree = self._mk_stats_tree(
            mid, "Legkelendőbb termékek", ("Termék", "Eladott db", "Bevétel"),
            (260, 110, 150), row=0, col=1)
        self.stats_low_tree = self._mk_stats_tree(
            mid, "Alacsony készlet (≤ 3 db)", ("Termék", "Vonalkód", "Készlet"),
            (300, 160, 100), row=1, col=0, colspan=2)

    def _mk_stats_tree(self, parent, title, cols, widths, row, col, colspan=1,
                       left=(0,)):
        wrap = tk.Frame(parent, bg=WINDOW_BG)
        wrap.grid(row=row, column=col, columnspan=colspan, sticky="nsew",
                  padx=6, pady=6)
        tk.Label(wrap, text=title, bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 13, "bold")).pack(anchor="w", pady=(0, 4))
        card = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER,
                        highlightthickness=1)
        card.pack(fill="both", expand=True)
        tv = ttk.Treeview(card, columns=cols, show="headings",
                          style="Admin.Treeview", selectmode="none")
        anchors = {i: "w" for i in left}
        for i, (cname, w) in enumerate(zip(cols, widths)):
            tv.heading(cname, text=cname)
            tv.column(cname, width=w, anchor=anchors.get(i, "e"))
        tv.tag_configure("even", background=ROW_ALT)
        tv.tag_configure("low", foreground=DANGER)
        tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(card, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        return tv

    def _stat_card(self, parent, caption, value, accent):
        p = Panel(parent, fill=SURFACE, outline=BORDER, radius=14,
                  parent_bg=WINDOW_BG, height=88, shadow=SHADOW)
        p.pack(side="left", fill="x", expand=True, padx=6)
        tk.Label(p, text=caption, bg=SURFACE, fg=MUTED,
                 font=(FONT_UI, 10, "bold")).pack(anchor="w", padx=16,
                                                  pady=(14, 0))
        tk.Label(p, text=value, bg=SURFACE, fg=accent,
                 font=(FONT_UI, 22, "bold")).pack(anchor="w", padx=16,
                                                  pady=(0, 12))

    def _refresh_stats(self):
        st = self.inv.compute_stats(low_threshold=3)

        for w in self.stats_cards.winfo_children():
            w.destroy()
        self._stat_card(self.stats_cards, "TERMÉKEK",
                        f"{st['product_count']} db", ACCENT)
        self._stat_card(self.stats_cards, "KÉSZLET ÖSSZESEN",
                        f"{st['total_units']} db", ACCENT)
        self._stat_card(self.stats_cards, "KÉSZLET ÉRTÉK (nettó)",
                        format_ft(st['total_value']), HEADER_BG)
        self._stat_card(self.stats_cards, "ÉVES ELADÁS",
                        f"{st['year_units']} db", SUCCESS)
        self._stat_card(self.stats_cards, "ÉVES BEVÉTEL",
                        format_ft(st['year_revenue']), SUCCESS_TEXT)

        mt = self.stats_month_tree
        mt.delete(*mt.get_children())
        for m in range(12):
            tags = ("even",) if m % 2 else ()
            mt.insert("", "end", values=(
                f"{m + 1}. {HU_MONTHS[m]}",
                st["monthly_units"][m],
                format_ft(st["monthly_revenue"][m])), tags=tags)

        tt = self.stats_top_tree
        tt.delete(*tt.get_children())
        for i, (p, sold, rev) in enumerate(st["top_products"][:25]):
            tags = ("even",) if i % 2 else ()
            tt.insert("", "end", values=(p.label, sold, format_ft(rev)),
                      tags=tags)

        lt = self.stats_low_tree
        lt.delete(*lt.get_children())
        for i, p in enumerate(st["low_stock"][:100]):
            tags = ["even"] if i % 2 else []
            tags.append("low")
            lt.insert("", "end", values=(p.label, p.barcode,
                                         int(p.stock or 0)), tags=tuple(tags))

    def _build_daily_tab(self):
        t = self.tab_daily
        self._daily_report = None

        bar = tk.Frame(t, bg=WINDOW_BG)
        bar.pack(fill="x", padx=8, pady=(10, 6))
        tk.Label(bar, text="Nap:", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 11)).pack(side="left")
        self.daily_day = ttk.Combobox(bar, width=16, state="readonly",
                                      font=(FONT_UI, 11), values=[])
        self.daily_day.pack(side="left", padx=(8, 16))
        self.daily_day.bind("<<ComboboxSelected>>",
                            lambda e: self._show_daily())

        RoundButton(bar, "💾  Zárás mentése…", fill=ACCENT, hover=ACCENT_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=186, height=40,
                    radius=10, font=(FONT_UI, 11, "bold"),
                    command=self._export_daily).pack(side="right", padx=(8, 0))
        RoundButton(bar, "🔄  Frissítés", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=WINDOW_BG,
                    width=130, height=40, radius=10, font=(FONT_UI, 11),
                    command=self._refresh_daily).pack(side="right")

        self.daily_cards = tk.Frame(t, bg=WINDOW_BG)
        self.daily_cards.pack(fill="x", padx=8, pady=(6, 6))

        mid = tk.Frame(t, bg=WINDOW_BG)
        mid.pack(fill="both", expand=True, padx=8, pady=(4, 10))
        mid.columnconfigure(0, weight=1, uniform="d")
        mid.columnconfigure(1, weight=1, uniform="d")
        mid.rowconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)

        self.daily_seller_tree = self._mk_stats_tree(
            mid, "Eladók bontása", ("Eladó", "Eladott db", "Bevétel"),
            (200, 120, 160), row=0, col=0)
        self.daily_prod_tree = self._mk_stats_tree(
            mid, "Aznap eladott termékek", ("Termék", "Méret", "Db", "Bevétel"),
            (230, 90, 70, 150), row=0, col=1, left=(0, 1))
        self.daily_wd_tree = self._mk_stats_tree(
            mid, "Aznapi elvitelek", ("Idő", "Termék", "Méret", "Db", "Indok"),
            (90, 250, 90, 60, 200), row=1, col=0, colspan=2, left=(0, 1, 2, 4))

    def _refresh_daily(self):
        """A nap-lista frissítése (friss diszk-olvasásból), majd megjelenítés."""
        days = self.inv.report_days()
        today = datetime.now().strftime("%Y.%m.%d")
        if today not in days:
            days = [today] + days
        cur = self.daily_day.get()
        self.daily_day.configure(values=days)
        if cur not in days:
            cur = days[0] if days else today
            self.daily_day.set(cur)
        self._show_daily()

    def _show_daily(self):
        day = self.daily_day.get() or datetime.now().strftime("%Y.%m.%d")
        rep = self.inv.daily_summary(day)
        self._daily_report = rep

        for w in self.daily_cards.winfo_children():
            w.destroy()
        self._stat_card(self.daily_cards, "BEVÉTEL",
                        format_ft(rep["total_revenue"]), SUCCESS_TEXT)
        self._stat_card(self.daily_cards, "NYUGTÁK",
                        f"{rep['receipts']} db", ACCENT)
        self._stat_card(self.daily_cards, "ELADOTT DB",
                        f"{rep['total_units']} db", ACCENT)
        self._stat_card(self.daily_cards, "ÁTLAG KOSÁR",
                        format_ft(rep["avg_basket"]), HEADER_BG)
        self._stat_card(self.daily_cards, "ELVITEL",
                        f"{rep['withdraw_units']} db",
                        DANGER if rep["withdraw_units"] else MUTED)

        st = self.daily_seller_tree
        st.delete(*st.get_children())
        for i, (seller, units, rev) in enumerate(rep["sellers"]):
            st.insert("", "end", values=(seller, units, format_ft(rev)),
                      tags=("even",) if i % 2 else ())

        pt = self.daily_prod_tree
        pt.delete(*pt.get_children())
        for i, (name, size, units, rev) in enumerate(rep["products"]):
            pt.insert("", "end", values=(name, size, units, format_ft(rev)),
                      tags=("even",) if i % 2 else ())

        wt = self.daily_wd_tree
        wt.delete(*wt.get_children())
        for i, r in enumerate(rep["withdrawals"]):
            stamp = as_text(r[0])
            clock = stamp[11:16] if len(stamp) >= 16 else ""
            wt.insert("", "end",
                      values=(clock, as_text(r[2]), as_text(r[3]),
                              parse_int(r[4]), as_text(r[5])),
                      tags=("even",) if i % 2 else ())

    def _daily_report_text(self, rep: dict) -> str:
        """Kinyomtatható/menthető, monospace-barát napi zárás szöveg."""
        W = 48
        L: list[str] = []
        L.append("=" * W)
        L.append("OXIGÉN CIPŐBOLT – NAPI ZÁRÁS".center(W))
        L.append(rep["day"].center(W))
        L.append("=" * W)
        L.append("")
        L.append(f"  Bevétel:        {format_ft(rep['total_revenue']):>27}")
        L.append(f"  Nyugták száma:  {str(rep['receipts']) + ' db':>27}")
        L.append(f"  Eladott:        {str(rep['total_units']) + ' db':>27}")
        L.append(f"  Átlag kosár:    {format_ft(rep['avg_basket']):>27}")
        if rep["withdraw_units"]:
            L.append(f"  Elvitel:        {str(rep['withdraw_units']) + ' db':>27}")
        L.append("")
        L.append("-" * W)
        L.append("  ELADÓK")
        L.append("-" * W)
        if rep["sellers"]:
            for seller, units, rev in rep["sellers"]:
                L.append(f"  {seller[:20]:<20}{str(units) + ' db':>7}"
                         f"{format_ft(rev):>17}")
        else:
            L.append("  (nincs eladás)")
        L.append("")
        L.append("-" * W)
        L.append("  ELADOTT TERMÉKEK")
        L.append("-" * W)
        if rep["products"]:
            for name, size, units, rev in rep["products"]:
                label = (f"{name} / {size}" if size else name)[:28]
                L.append(f"  {label:<28}{str(units) + 'db':>5}"
                         f"{format_ft(rev):>11}")
        else:
            L.append("  (nincs eladás)")
        if rep["withdrawals"]:
            L.append("")
            L.append("-" * W)
            L.append("  ELVITELEK")
            L.append("-" * W)
            for r in rep["withdrawals"]:
                nm, sz = as_text(r[2]), as_text(r[3])
                label = (f"{nm} / {sz}" if sz else nm)[:26]
                L.append(f"  {label:<26}{str(parse_int(r[4])) + 'db':>5}"
                         f"  {as_text(r[5])[:13]}")
        L.append("")
        L.append("=" * W)
        L.append(f"Készült: {datetime.now().strftime('%Y.%m.%d %H:%M')}".center(W))
        L.append("=" * W)
        return "\n".join(L) + "\n"

    def _export_daily(self):
        rep = self._daily_report
        if not rep:
            return
        if not rep["receipts"] and not rep["withdrawals"]:
            messagebox.showinfo(
                "Napi zárás",
                f"Ezen a napon ({rep['day']}) nincs rögzített eladás vagy "
                "elvitel, nincs mit menteni.", parent=self)
            return
        default = f"napi_zaras_{rep['day'].replace('.', '-').strip('-')}.txt"
        path = filedialog.asksaveasfilename(
            parent=self, title="Napi zárás mentése",
            defaultextension=".txt", initialfile=default,
            filetypes=[("Szövegfájl", "*.txt"), ("Minden fájl", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self._daily_report_text(rep))
        except Exception as exc:
            messagebox.showerror("Hiba", f"Nem sikerült menteni:\n{exc}",
                                 parent=self)
            return
        messagebox.showinfo("Kész", f"Napi zárás elmentve:\n{path}",
                            parent=self)

    def _build_log_tab(self):
        t = self.tab_log
        bar = tk.Frame(t, bg=WINDOW_BG)
        bar.pack(fill="x", padx=8, pady=(10, 6))
        tk.Label(bar, text="Napló típusa:", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 11)).pack(side="left")
        self.log_kind = ttk.Combobox(bar, width=24, state="readonly",
                                     font=(FONT_UI, 11),
                                     values=["Eladások", "Elvitelek"])
        self.log_kind.current(0)
        self.log_kind.pack(side="left", padx=(8, 16))
        self.log_kind.bind("<<ComboboxSelected>>",
                           lambda e: self._refresh_log())
        tk.Label(bar, text="Keresés:", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 11)).pack(side="left")
        self.log_search = tk.Entry(bar, font=(FONT_UI, 12), width=24,
                                   relief="flat", bg=SURFACE, fg=TEXT,
                                   insertbackground=ACCENT)
        self.log_search.pack(side="left", padx=(8, 0), ipady=4)
        self.log_search.bind("<KeyRelease>", lambda e: self._refresh_log())
        RoundButton(bar, "🔄  Frissítés", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=WINDOW_BG,
                    width=130, height=40, radius=10, font=(FONT_UI, 11),
                    command=self._refresh_log).pack(side="right")
        self.log_summary = tk.Label(bar, text="", bg=WINDOW_BG, fg=MUTED,
                                    font=(FONT_UI, 10))
        self.log_summary.pack(side="right", padx=12)

        self.log_card = tk.Frame(t, bg=SURFACE, highlightbackground=BORDER,
                                 highlightthickness=1)
        self.log_card.pack(fill="both", expand=True, padx=8, pady=(4, 10))
        self.log_tree = None

    def _refresh_log(self):
        kind = self.log_kind.get()
        if kind == "Eladások":
            header = SALES_LOG_HEADER
            widths = (130, 120, 120, 220, 90, 60, 90, 90, 60)
            rows = self.inv.read_sales_log_rows()
        else:
            header = ["Dátum", "Vonalkód", "Megnevezés", "Méret",
                      "Darab", "Indok", "Hónap"]
            widths = (140, 130, 240, 90, 70, 180, 70)
            rows = self.inv.read_log_rows()

        for w in self.log_card.winfo_children():
            w.destroy()
        tv = ttk.Treeview(self.log_card, columns=header, show="headings",
                          style="Admin.Treeview", selectmode="browse")
        for cname, w in zip(header, widths):
            tv.heading(cname, text=cname)
            anc = "e" if cname in ("Darab", "Egységár", "Összeg", "Hónap") else "w"
            tv.column(cname, width=w, anchor=anc)
        tv.tag_configure("even", background=ROW_ALT)
        tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(self.log_card, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_tree = tv

        q = self.log_search.get().strip().lower()
        shown = 0
        for i, r in enumerate(reversed(rows)):
            vals = ["" if v is None else v for v in r]
            if q and q not in " ".join(str(v) for v in vals).lower():
                continue
            tv.insert("", "end", values=vals,
                      tags=("even",) if shown % 2 else ())
            shown += 1
        self.log_summary.configure(
            text=f"{shown} bejegyzés" if not q else f"{shown} találat")

    def _build_staff_tab(self):
        t = self.tab_staff
        wrap = tk.Frame(t, bg=WINDOW_BG)
        wrap.pack(fill="both", expand=True, padx=16, pady=14)
        wrap.columnconfigure(0, weight=1)
        wrap.columnconfigure(1, weight=1)

        left = tk.Frame(wrap, bg=WINDOW_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        tk.Label(left, text="Dolgozók (eladók)", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 14, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(left, text="Belépéskor a dolgozó kiválasztja a nevét; ez "
                            "kerül az eladás-naplóba.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10),
                 wraplength=380, justify="left").pack(anchor="w", pady=(0, 8))
        lb_card = tk.Frame(left, bg=SURFACE, highlightbackground=BORDER,
                           highlightthickness=1)
        lb_card.pack(fill="both", expand=True)
        self.staff_list = tk.Listbox(lb_card, font=(FONT_UI, 13),
                                     activestyle="none", relief="flat",
                                     bg=SURFACE, fg=TEXT,
                                     selectbackground=SELECTION,
                                     selectforeground=TEXT, height=10)
        self.staff_list.pack(side="left", fill="both", expand=True, padx=4,
                             pady=4)
        sb = ttk.Scrollbar(lb_card, orient="vertical",
                           command=self.staff_list.yview)
        self.staff_list.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        add_bar = tk.Frame(left, bg=WINDOW_BG)
        add_bar.pack(fill="x", pady=(10, 0))
        self.staff_entry = tk.Entry(add_bar, font=(FONT_UI, 13), relief="flat",
                                    bg=SURFACE, fg=TEXT, insertbackground=ACCENT)
        self.staff_entry.pack(side="left", fill="x", expand=True, ipady=5)
        self.staff_entry.bind("<Return>", lambda e: self._add_staff())
        RoundButton(add_bar, "➕  Hozzáad", fill=SUCCESS, hover=SUCCESS_HOVER,
                    fg="white", parent_bg=WINDOW_BG, width=130, height=40,
                    radius=10, font=(FONT_UI, 12, "bold"),
                    command=self._add_staff).pack(side="left", padx=(8, 0))
        RoundButton(left, "🗑  Kijelölt törlése", fill=DANGER,
                    hover=DANGER_HOVER, fg="white", parent_bg=WINDOW_BG,
                    width=200, height=40, radius=10, font=(FONT_UI, 12, "bold"),
                    command=self._del_staff).pack(anchor="w", pady=(10, 0))

        right = tk.Frame(wrap, bg=WINDOW_BG)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        tk.Label(right, text="Admin jelszó módosítása", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 14, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(right, text="Az admin felület belépési jelszava. "
                             "Alap: „admin”.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10),
                 wraplength=380, justify="left").pack(anchor="w", pady=(0, 12))
        pwcard = Panel(right, fill=SURFACE, outline=BORDER, radius=14,
                       parent_bg=WINDOW_BG, shadow=SHADOW)
        pwcard.pack(fill="x")
        inner = tk.Frame(pwcard, bg=SURFACE)
        inner.pack(fill="x", padx=18, pady=16)
        tk.Label(inner, text="Új jelszó", bg=SURFACE, fg=MUTED,
                 font=(FONT_UI, 11)).grid(row=0, column=0, sticky="e",
                                          padx=(0, 8), pady=6)
        self.pw1 = tk.Entry(inner, show="•", font=(FONT_UI, 13), width=22,
                            relief="flat", bg=WINDOW_BG, fg=TEXT,
                            insertbackground=ACCENT)
        self.pw1.grid(row=0, column=1, pady=6, ipady=4)
        tk.Label(inner, text="Mégegyszer", bg=SURFACE, fg=MUTED,
                 font=(FONT_UI, 11)).grid(row=1, column=0, sticky="e",
                                          padx=(0, 8), pady=6)
        self.pw2 = tk.Entry(inner, show="•", font=(FONT_UI, 13), width=22,
                            relief="flat", bg=WINDOW_BG, fg=TEXT,
                            insertbackground=ACCENT)
        self.pw2.grid(row=1, column=1, pady=6, ipady=4)
        RoundButton(inner, "🔒  Jelszó mentése", fill=ACCENT, hover=ACCENT_HOVER,
                    fg="white", parent_bg=SURFACE, width=200, height=42,
                    radius=12, font=(FONT_UI, 12, "bold"),
                    command=self._change_password).grid(
            row=2, column=0, columnspan=2, pady=(14, 0))

    def _refresh_staff(self):
        self.staff_list.delete(0, "end")
        for n in get_employees():
            self.staff_list.insert("end", n)

    def _add_staff(self):
        name = self.staff_entry.get().strip()
        if not name:
            return
        names = get_employees()
        if name in names:
            messagebox.showinfo("Már létezik", f"„{name}” már szerepel.",
                                parent=self)
            return
        names.append(name)
        set_employees(names)
        self.staff_entry.delete(0, "end")
        self._refresh_staff()

    def _del_staff(self):
        sel = self.staff_list.curselection()
        if not sel:
            messagebox.showinfo("Nincs kijelölve",
                                "Előbb válassz ki egy dolgozót.", parent=self)
            return
        name = self.staff_list.get(sel[0])
        if not messagebox.askyesno("Törlés",
                                   f"Törlöd „{name}” dolgozót a listából?",
                                   parent=self):
            return
        names = [n for n in get_employees() if n != name]
        set_employees(names)
        self._refresh_staff()

    def _change_password(self):
        p1, p2 = self.pw1.get(), self.pw2.get()
        if not p1:
            messagebox.showwarning("Üres jelszó",
                                   "Adj meg egy jelszót.", parent=self)
            return
        if p1 != p2:
            messagebox.showwarning("Eltérés",
                                   "A két jelszó nem egyezik.", parent=self)
            return
        set_admin_password(p1)
        self.pw1.delete(0, "end")
        self.pw2.delete(0, "end")
        messagebox.showinfo("Kész", "Az admin jelszó megváltozott.",
                            parent=self)

    # ------------------------------------------------------------ Frissítés --

    def _build_update_tab(self):
        t = self.tab_update
        wrap = tk.Frame(t, bg=WINDOW_BG)
        wrap.pack(fill="both", expand=True, padx=18, pady=16)

        tk.Label(wrap, text="Programfrissítés", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 16, "bold")).pack(anchor="w")
        tk.Label(wrap, text="A program indításkor csendben megnézi, van-e újabb "
                            "változat. A dolgozók ebből semmit nem látnak; a "
                            "frissítést innen lehet telepíteni.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10),
                 wraplength=760, justify="left").pack(anchor="w", pady=(2, 14))

        card = Panel(wrap, fill=SURFACE, outline=BORDER, radius=14,
                     parent_bg=WINDOW_BG, shadow=SHADOW)
        card.pack(fill="x")
        inner = tk.Frame(card, bg=SURFACE)
        inner.pack(fill="x", padx=20, pady=18)

        def sor(cimke, ertek, vastag=False):
            r = tk.Frame(inner, bg=SURFACE)
            r.pack(fill="x", pady=3)
            tk.Label(r, text=cimke, bg=SURFACE, fg=MUTED, width=18,
                     anchor="w", font=(FONT_UI, 11)).pack(side="left")
            lbl = tk.Label(r, text=ertek, bg=SURFACE, fg=TEXT, anchor="w",
                           font=(FONT_UI, 12, "bold" if vastag else "normal"))
            lbl.pack(side="left")
            return lbl

        self.upd_version = sor("Jelenlegi verzió:", f"v{APP_VERSION}", True)
        self.upd_repo = sor("Frissítési forrás:", "—")
        self.upd_status = sor("Állapot:", "—", True)

        gombok = tk.Frame(inner, bg=SURFACE)
        gombok.pack(anchor="w", pady=(16, 0))
        self.upd_check_btn = RoundButton(
            gombok, "🔄  Frissítés keresése", fill=NEUTRAL_BTN,
            hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=SURFACE, width=210,
            height=42, radius=12, font=(FONT_UI, 12, "bold"),
            command=self._check_update)
        self.upd_check_btn.pack(side="left")
        self.upd_install_btn = RoundButton(
            gombok, "⬆  Telepítés most", fill=SUCCESS, hover=SUCCESS_HOVER,
            fg="white", parent_bg=SURFACE, width=190, height=42, radius=12,
            font=(FONT_UI, 12, "bold"), command=self._install_update)
        self.upd_install_btn.pack(side="left", padx=(10, 0))
        RoundButton(gombok, "⚙  Tároló beállítása", fill=NEUTRAL_BTN,
                    hover=NEUTRAL_BTN_HOVER, fg=TEXT, parent_bg=SURFACE,
                    width=195, height=42, radius=12, font=(FONT_UI, 12),
                    command=self._ask_repo).pack(side="left", padx=(10, 0))

        self.upd_progress = tk.Label(inner, text="", bg=SURFACE, fg=ACCENT,
                                     font=(FONT_UI, 11), anchor="w")
        self.upd_progress.pack(fill="x", pady=(12, 0))

        tk.Label(wrap, text="Mi újság az új verzióban", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 12, "bold")).pack(anchor="w", pady=(18, 4))
        notes_card = tk.Frame(wrap, bg=SURFACE, highlightbackground=BORDER,
                              highlightthickness=1)
        notes_card.pack(fill="both", expand=True)
        self.upd_notes = tk.Text(notes_card, font=(FONT_UI, 11), relief="flat",
                                 bg=SURFACE, fg=TEXT, wrap="word", height=9,
                                 padx=12, pady=10)
        self.upd_notes.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(notes_card, orient="vertical",
                           command=self.upd_notes.yview)
        self.upd_notes.configure(yscrollcommand=sb.set, state="disabled")
        sb.pack(side="right", fill="y")

        tk.Label(wrap, text="A frissítés a program fájljait cseréli le. Az "
                            "Excel-táblát, a beállításokat és a mentéseket "
                            "SOHA nem írja felül, és telepítés előtt mindenről "
                            "biztonsági másolat készül a backup mappába.",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 9), wraplength=760,
                 justify="left").pack(anchor="w", pady=(10, 0))

    def _set_notes(self, szoveg: str):
        self.upd_notes.configure(state="normal")
        self.upd_notes.delete("1.0", "end")
        self.upd_notes.insert("1.0", szoveg or "—")
        self.upd_notes.configure(state="disabled")

    def _upd_progress(self, szoveg: str):
        """Szálbiztos állapotkiírás (a háttérszál ezen keresztül üzen)."""
        try:
            self.after(0, lambda: self.upd_progress.configure(text=szoveg))
        except tk.TclError:
            pass

    def _refresh_update(self):
        if frissito is None:
            self.upd_repo.configure(text="a frissítő modul hiányzik "
                                         "(frissito.py)")
            self.upd_status.configure(text="nem elérhető", fg=MUTED)
            return

        repo = frissito.get_repo()
        self.upd_repo.configure(text=repo or "nincs beállítva")

        allapot = FRISSITES.get("allapot", "nincs")
        info = FRISSITES.get("info")
        if not repo:
            self.upd_status.configure(text="kikapcsolva", fg=MUTED)
            self._set_notes("Állítsd be a GitHub-tárolót a „Tároló beállítása” "
                            "gombbal, és a program ezentúl magától figyeli az "
                            "új verziókat.")
        elif allapot == "van" and info is not None:
            self.upd_status.configure(text=f"új verzió: v{info.version}",
                                      fg=SUCCESS_TEXT)
            self._set_notes(info.notes or "(a kiadáshoz nincs leírás)")
            try:
                self._nb.tab(self._update_tab_index, text="  ⬆  Frissítés •  ")
            except Exception:
                pass
        elif allapot == "naprakesz":
            self.upd_status.configure(text="a program naprakész", fg=TEXT)
            self._set_notes("Nincs újabb kiadás.")
        elif allapot == "keres":
            self.upd_status.configure(text="ellenőrzés folyamatban…", fg=MUTED)
        else:
            self.upd_status.configure(text="még nem ellenőriztük", fg=MUTED)

    def _check_update(self):
        if frissito is None:
            return
        if not frissito.get_repo():
            messagebox.showinfo("Nincs beállítva",
                                "Előbb add meg a GitHub-tárolót a „Tároló "
                                "beállítása” gombbal.", parent=self)
            return
        self.upd_progress.configure(text="Keresés…")

        def munka():
            try:
                info = frissito.check(timeout=15)
                FRISSITES["info"] = info
                FRISSITES["allapot"] = "van" if info else "naprakesz"
                self._upd_progress("")
            except Exception as exc:
                FRISSITES["allapot"] = "hiba"
                self._upd_progress("")
                uzenet = str(exc)
                try:
                    self.after(0, lambda: messagebox.showerror(
                        "Nem sikerült", uzenet, parent=self))
                except tk.TclError:
                    return
            try:
                self.after(0, self._refresh_update)
            except tk.TclError:
                pass

        threading.Thread(target=munka, daemon=True).start()

    def _install_update(self):
        info = FRISSITES.get("info")
        if frissito is None or info is None:
            messagebox.showinfo("Nincs mit telepíteni",
                                "Előbb keress frissítést.", parent=self)
            return
        if not messagebox.askyesno(
                "Frissítés telepítése",
                f"Jelenlegi verzió:  v{APP_VERSION}\n"
                f"Új verzió:         v{info.version}\n\n"
                "A program letölti és telepíti az új változatot, majd "
                "újraindul.\n\n"
                "Az Excel-tábla és a beállítások érintetlenek maradnak, és "
                "minden cserélt fájlról mentés készül.\n\n"
                "Most ne legyen folyamatban eladás. Folytatod?",
                parent=self):
            return

        self.upd_progress.configure(text="Indul…")

        def munka():
            try:
                backup = frissito.update_now(info, progress=self._upd_progress)
            except Exception as exc:
                uzenet = str(exc)
                self._upd_progress("")
                try:
                    self.after(0, lambda: messagebox.showerror(
                        "A frissítés nem sikerült", uzenet, parent=self))
                except tk.TclError:
                    pass
                return

            def kesz():
                messagebox.showinfo(
                    "Kész",
                    f"A program frissült a v{info.version} verzióra, és most "
                    "újraindul.\n\n"
                    + (f"A régi fájlok mentése:\n{backup}" if backup else ""),
                    parent=self)
                self.destroy()
                frissito.restart()

            try:
                self.after(0, kesz)
            except tk.TclError:
                pass

        threading.Thread(target=munka, daemon=True).start()

    def _ask_repo(self):
        """A GitHub-tároló megadása (tulajdonos/tároló alakban)."""
        if frissito is None:
            return
        dlg = tk.Toplevel(self)
        dlg.title("Frissítési forrás")
        dlg.configure(bg=WINDOW_BG)
        dlg.transient(self)
        dlg.resizable(False, False)

        tk.Label(dlg, text="⚙  Frissítési forrás", bg=WINDOW_BG, fg=TEXT,
                 font=(FONT_UI, 16, "bold")).pack(padx=30, pady=(22, 4))
        tk.Label(dlg, text="Annak a GitHub-tárolónak a neve, ahova a program új "
                           "kiadásait töltöd fel.\nPélda:  "
                           "buczibotond/oxigen-keszletezo",
                 bg=WINDOW_BG, fg=MUTED, font=(FONT_UI, 10),
                 justify="left").pack(padx=30, pady=(0, 14))
        ent = tk.Entry(dlg, font=(FONT_UI, 13), justify="center", relief="flat",
                       bg=SURFACE, fg=TEXT, insertbackground=ACCENT, width=34)
        ent.insert(0, frissito.get_repo())
        ent.pack(padx=30, ipady=6)
        ent.focus_set()

        def ment():
            frissito.set_repo(ent.get().strip())
            dlg.destroy()
            self._refresh_update()
            self._check_update()

        bar = tk.Frame(dlg, bg=WINDOW_BG)
        bar.pack(padx=30, pady=(18, 22))
        RoundButton(bar, "Mentés", fill=ACCENT, hover=ACCENT_HOVER, fg="white",
                    parent_bg=WINDOW_BG, width=150, height=44, radius=12,
                    font=(FONT_UI, 13, "bold"),
                    command=ment).pack(side="left", padx=6)
        RoundButton(bar, "Mégse", fill=NEUTRAL_BTN, hover=NEUTRAL_BTN_HOVER,
                    fg=TEXT, parent_bg=WINDOW_BG, width=120, height=44,
                    radius=12, font=(FONT_UI, 12),
                    command=dlg.destroy).pack(side="left", padx=6)
        ent.bind("<Return>", lambda e: ment())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_window(dlg, 460, 270)
        dlg.grab_set()
        self.wait_window(dlg)

def enable_dpi_awareness() -> None:
    """Windowson éles (nem elmosódott) megjelenés magas felbontású kijelzőn.

    A hívásnak az első Tk-ablak létrehozása ELŐTT kell megtörténnie.
    """
    if sys.platform != "win32":
        return
    try:
        from ctypes import windll
        try:
            windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            windll.user32.SetProcessDPIAware()
    except Exception:
        pass

def dev_requested() -> bool:
    """Fejlesztői mód kérése: --dev / -d kapcsoló, vagy KESZLET_DEV környezeti
    változó (1/true/yes/on/igen). Az appon belül Ctrl+Shift+D-vel is váltható."""
    argv = sys.argv[1:]
    if "--dev" in argv or "-d" in argv:
        return True
    val = os.environ.get("KESZLET_DEV", "").strip().lower()
    return val in ("1", "true", "yes", "on", "igen")

def main():
    enable_dpi_awareness()
    root = tk.Tk()
    root.withdraw()
    _select_fonts(root)
    start_update_check()          # csendben, háttérben – a kasszát nem érinti
    path = resolve_workbook(parent=root)
    if path is None:
        messagebox.showinfo("Kilépés", "Nem választottál Excel-fájlt. Kilépés.")
        root.destroy()
        return
    try:
        inventory = Inventory(path)
    except Exception as exc:
        messagebox.showerror(
            "Hiba az Excel megnyitásakor",
            f"Nem sikerült beolvasni a fájlt:\n{path}\n\n{exc}")
        root.destroy()
        return
    root.destroy()

    dev = dev_requested()
    while True:
        login = LoginWindow()
        login.mainloop()
        choice = login.result
        if not choice:
            break
        role, operator = choice

        try:
            inventory.load()
        except Exception:
            pass

        if role == "admin":
            win = AdminWindow(inventory)
        else:
            win = App(inventory, dev=dev, operator=operator)
        win.mainloop()

        if not getattr(win, "relogin", False):
            break

if __name__ == "__main__":
    main()
