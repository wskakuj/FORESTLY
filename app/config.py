"""
Forestly — Konfiguracja globalna
==========================================
Zależności: brak (moduł bazowy)
Odpowiada za: wszystkie stałe, kolory, ścieżki, dane terytorialne,
              funkcje pomocnicze (marginesy, historia, kolejność PDF).

Ten moduł jest importowany przez praktycznie każdy inny moduł.
Nie importuje niczego z projektu — tylko biblioteki standardowe i zewnętrzne.
"""

import os
import re
import sys
import json
from pathlib import Path

import customtkinter as ctk

# --- WERSJA I AKTUALIZACJA ---
CURRENT_VERSION = "v2.0.68"
GITHUB_USER = "wskakuj"
GITHUB_REPO = "FORESTLY"

# --- KONFIGURACJA GUI ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

ENCODING = "cp852"
ORDER_FILE_NAME = "pdf_merge_orders.json"

# --- ŚCIEŻKI PLIKÓW KONFIGURACYJNYCH ---
def _get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent  # cofamy się z app/ do głównego katalogu

APP_DIR = _get_app_dir()
CONFIG_DIR = APP_DIR / "config"
TERRITORY_DATA_FILE = CONFIG_DIR / "territory.json"
HISTORY_FILE = APP_DIR / "folder_history.json"
MARGINS_FILE = APP_DIR / "margins_config.json"
SETTINGS_FILE = APP_DIR / "settings.json"

# --- PALETA KOLORÓW ---
# Wszystkie kolory używane w GUI, w jednym miejscu.
# Ułatwia zmianę motywu i ewentualne dodanie trybu jasnego.
COLORS = {
    "primary":        "#0078D7",
    "primary_hover":  "#005A9E",
    "primary_dark":   "#0067C0",
    "success":        "#27ae60",
    "success_hover":  "#219653",
    "danger":         "#DC143C",
    "danger_dark":    "#8B0000",
    "danger_hover":   "#A52A2A",
    "warning":        "#D83B01",
    "bg_dark":        "#1E1E1E",
    "bg_card":        "#252526",
    "bg_card_alt":    "#333333",
    "bg_card_hover":  "#444444",
    "text_primary":   "#E0E0E0",
    "text_muted":     "#888888",
    "text_dim":       "#A0A0A0",
    "text_log":       "#D4D4D4",
    "border":         "#333333",
    "icon_folder":    "#FFD700",
    "icon_start":     "#00FA9A",
    "icon_stop":      "#DC143C",
    "arrow":          "#555555",
    "pending":        "#555555",
}

# --- SEKWENCJE DO USUWANIA Z TXT ---
SEQUENCES_TO_REMOVE = [
    b"\x1b(s16.67H\x1b&l5E\x1b&a20L",   # nagłówek PCL WSKAZ1 (agencja)
    b"\x1b(s16.67H\x1b&l4E\x1b&a1L",
    b"\x1b(s16.67H\x1b&l5E\x1b&a9L",
    b"\x1b(s16.67H\x1b&l5E\x1b&a14L",
    b"\x1b(s16.67H\x1b&l9E\x1b&a8L",
    b"\x1b&l6E\x1b&a10L\x1b(s3T",
    b"\x1b(s16.67H\x1b&l5E\x1b&a0L",
    b"\x1b(s16.67H\x1b&l9E\x1b&a8L ",
    b"\x1b(s16.67H\x1b&l9E\x1b&a8L\xa0",
]

MACRO_MAP = {"optax": "OPTX", "tab_klw3": "TAB_KLW3", "wskaz1": "WSKAZ1"}

FILTER_ALIASES = {
    "WSZYSTKIE": {"WSZYSTKIE"},
    "Wszystkie": {"WSZYSTKIE"},
    "REJESTR1": {"REJESTR1"},
    "OPTAX": {"OPTAX"},
    "TAB_KLW3": {"TAB_KLW3"},
    "WSKAZ1": {"WSKAZ1", "WSK_ZB"},
    "WSK_ZB": {"WSKAZ1", "WSK_ZB"},
    "HALIZNY": {"HALIZNY"},
    "WYK_NEG": {"WYK_NEG", "WYKNEG"},
    "WYKNEG": {"WYK_NEG", "WYKNEG"},
    "OPIS": {"OPIS"},
    "ZEST1": {"ZEST1"},
    "WK_ZM1": {"WK_ZM1"},
}

# prefiks numeru pozycji w ustawionym układzie PDF (np. "03_OPTAX.pdf")
RE_PREFIX_UKLADU = re.compile(r"^\d{2}_")

PDF_ORDER_TEMPLATES = [
    {"key": "TITLE", "label": "Strona Tytułowa", "aliases": ["upul", "str_tyt", "strtyt"]},
    {"key": "OPIS", "label": "Opis ogólny", "aliases": ["opis", "op_ogplan"]},
    {"key": "TAB_KLW3", "label": "Zestawienie powierzchni gruntów i miąższości drzewostanu wg gatunków panujących (głównych) wg klas i podklas wieku", "aliases": ["tab_klw3.pdf"]},
    {"key": "OPTAX", "label": "Opisy taksacyjne lasu i gruntów przeznaczonych do zalesienia", "aliases": ["optax.pdf"]},
    {"key": "WSK_ZB", "label": "Zestawienie czynności gospodarczych projektowanych na 10 lat", "aliases": ["wsk_zb.pdf", "wsk_zb"]},
    {"key": "WYK_NEG", "label": "Wykaz d-stanów do przebudowy (negatywnych i źle produkujących)", "aliases": ["wyk_neg.pdf"]},
    {"key": "HALIZNY", "label": "Zestawienie pow. leśnych nie zalesionych", "aliases": ["halizny.pdf"]},
    {"key": "REJESTR1", "label": "Rejestr działek leśnych i wskazania gospodarcze w zakresie gospodarki leśnej", "aliases": ["rejestr1.pdf"]},
    {"key": "ZEST1", "label": "Skorowidz działek leśnych", "aliases": ["zest1.pdf", "skorowidz dz"]},
    {"key": "WK_ZM1", "label": "Wykaz rozbieżności między ewidencją gruntów a stanem faktycznym", "aliases": ["wk_zm1.pdf"]},
    {"key": "WSKAZ1", "label": "Wskazówki gospodarki leśnej", "aliases": ["wskaz1.pdf", "wskaz1"]},
    {"key": "SKROTY", "label": "Wykaz skrótów i symboli", "aliases": ["skroty"]},
    # mapa dołączana na końcu pakietu (plik *.pdf z "mapa" w nazwie,
    # np. "mapa.pdf" albo "CHORZEWO_mapa.pdf")
    {"key": "MAPA", "label": "Mapa", "aliases": ["mapa"]},
]

EXCEL_SHEET_DEFAULTS = [
    ("Zestawienie", 8, 9),
    ("WykazPow", 7, 9),
    ("OT", 10, 9),
    ("WykazWlasc", 7, 9),
    ("WykazDzialek", 7, 9),
    ("Skroty", 6, 9),
    ("REJ", 10, 9),
    ("TPM_FL", 5, 9),
    ("TPM_TH", 5, 9),
]


# --- DANE TERYTORIALNE ---
def load_territory_data() -> dict:
    candidates = [TERRITORY_DATA_FILE]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / "config" / "territory.json")
    try:
        candidates.append(Path(__file__).resolve().parent.parent / "config" / "territory.json")
    except Exception:
        pass
    for path in candidates:
        try:
            if path.exists():
                data = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(data, dict) and data:
                    if path != TERRITORY_DATA_FILE:
                        try:
                            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
                            TERRITORY_DATA_FILE.write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
                        except Exception:
                            pass
                    return data
        except Exception as e:
            print(f"[INFO] Błąd wczytania danych terytorialnych z {path}: {e}")
    print("[INFO] Brak poprawnego pliku config/territory.json. Listy terytorialne będą puste.")
    return {}


TERRITORY_DATA = load_territory_data()


# --- FUNKCJE POMOCNICZE ---
def kill_orphan_office_processes():
    import subprocess
    try:
        cmd = 'powershell "Get-Process -Name WINWORD, EXCEL -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -eq 0} | Stop-Process -Force"'
        subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f"[INFO] Błąd czyszczenia procesów tła: {e}")


def load_margins():
    if MARGINS_FILE.exists():
        try:
            return json.loads(MARGINS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_margins(data):
    try:
        MARGINS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[INFO] Błąd zapisu marginesów: {e}")


def clean_xml_incompatible(text):
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def flatten_rel_path(rel_path):
    parts = rel_path.parts
    new_parts = [p for p in parts[:-1] if not p.upper().startswith("WOL")]
    new_parts.append(parts[-1])
    return Path(*new_parts)


def normalize_name(name: str) -> str:
    return name.strip().lower()


def template_matches(template, pdf_name: str) -> bool:
    # pomijamy prefiks pozycji w układzie (np. "04_OPTAX.pdf" -> "OPTAX.pdf") —
    # dodawany przez scalanie, nie może psuć rozpoznawania przy ponownym biegu
    name = normalize_name(RE_PREFIX_UKLADU.sub("", pdf_name))
    for alias in template["aliases"]:
        alias = normalize_name(alias)
        if alias.endswith(".pdf"):
            if name == alias:
                return True
        else:
            if alias in name:
                return True
    return False


def get_default_template_keys():
    return [t["key"] for t in PDF_ORDER_TEMPLATES]


def get_order_store_path(folder: Path) -> Path:
    return folder / ORDER_FILE_NAME


def load_order_store(folder: Path):
    store = get_order_store_path(folder)
    if not store.exists():
        return {}
    try:
        return json.loads(store.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_order_store(folder: Path, data: dict):
    store = get_order_store_path(folder)
    store.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def get_saved_template_order(folder: Path, mode_key: str):
    data = load_order_store(folder)
    saved = data.get(mode_key)
    if isinstance(saved, list) and saved:
        valid = [t["key"] for t in PDF_ORDER_TEMPLATES]
        saved = [k for k in saved if k in valid]
        # dopisanie nowych kluczy (np. WSKAZ1) na ich domyślnych pozycjach,
        # z zachowaniem własnej kolejności starego zapisu
        missing = [k for k in get_default_template_keys() if k not in saved]
        # brakujące pozycje (np. dopisane w nowszej wersji programu, a nieobecne
        # w starym zapisie) dokładamy NA KOŃCU — własna kolejność użytkownika
        # zostaje nietknięta i nic nie wędruje na początek zestawienia
        for mk in missing:
            saved.append(mk)
        return saved
    return get_default_template_keys()


def set_saved_template_order(folder: Path, mode_key: str, order_keys):
    data = load_order_store(folder)
    data[mode_key] = order_keys
    save_order_store(folder, data)


def build_ordered_pdfs_from_templates(pdfs, template_keys, excluded=None):
    """Porządkuje PDF-y wg kluczy szablonów.

    excluded — lista kluczy szablonów wykluczonych z scalania:
    pasujące do nich pliki NIE trafiają do wyniku (ani do 'zaległości'
    dokładanych na końcu).
    """
    excluded = set(excluded or [])
    ordered = []
    used = set()
    template_map = {t["key"]: t for t in PDF_ORDER_TEMPLATES}
    excluded_templates = [template_map[k] for k in excluded if k in template_map]
    for key in template_keys:
        if key in excluded:
            continue
        template = template_map.get(key)
        if not template:
            continue
        matches = [p for p in pdfs if p not in used and template_matches(template, p.name)]
        for p in matches:
            ordered.append(p)
            used.add(p)
    for p in pdfs:
        if p in used:
            continue
        if any(template_matches(t, p.name) for t in excluded_templates):
            continue  # wykluczony — nie dokładamy go na końcu
        ordered.append(p)

    # 'Skróty i symbole' — ZAWSZE na samym końcu scalonego PDF, także gdy
    # w folderze są pliki niepasujące do żadnego szablonu (dokładane wyżej).
    skroty_tpl = template_map.get("SKROTY")
    if skroty_tpl is not None and "SKROTY" not in excluded:
        ordered.sort(key=lambda _p: 1 if template_matches(skroty_tpl, _p.name) else 0)
    return ordered


def get_saved_excluded_templates(folder: Path, mode_key: str):
    """Klucze szablonów wykluczone z scalania dla danego trybu."""
    data = load_order_store(folder)
    saved = data.get(mode_key + "_excluded")
    if isinstance(saved, list):
        valid = {t["key"] for t in PDF_ORDER_TEMPLATES}
        return [k for k in saved if k in valid]
    return []


def set_saved_excluded_templates(folder: Path, mode_key: str, excluded_keys):
    data = load_order_store(folder)
    valid = [t["key"] for t in PDF_ORDER_TEMPLATES]
    data[mode_key + "_excluded"] = [k for k in (excluded_keys or []) if k in valid]
    save_order_store(folder, data)


def is_file_locked(filepath):
    filepath = Path(filepath)
    if not filepath.exists():
        return False
    try:
        with open(filepath, "a"):
            pass
    except PermissionError:
        return True
    except Exception:
        pass
    return False


def normalize_filter_selection(file_filter):
    if file_filter is None:
        return {"WSZYSTKIE"}
    if isinstance(file_filter, str):
        values = [file_filter]
    elif isinstance(file_filter, (list, tuple, set)):
        values = list(file_filter)
    else:
        values = [str(file_filter)]
    normalized = set()
    for value in values:
        item = str(value).strip().upper()
        if not item:
            continue
        normalized.update(FILTER_ALIASES.get(item, {item}))
    if not normalized or "WSZYSTKIE" in normalized:
        return {"WSZYSTKIE"}
    return normalized


def add_tooltip(widget, text):
    try:
        from CTkToolTip import CTkToolTip
        CTkToolTip(widget, message=text, delay=0.5)
    except ImportError:
        pass
