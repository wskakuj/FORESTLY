# -*- coding: utf-8 -*-
"""Generowanie plików .VAL (rozliczenie geodezyjne) z map GEO-MAP (.MAP).

Plik .VAL zawiera, dla każdej DZIAŁKI, wydzielenia na niej leżące wraz z ich
powierzchniami. Powierzchnia to CZĘŚĆ WSPÓLNA (przecięcie) poligonu wydzielenia
i poligonu działki — tak liczy to GEO-MAP.

W mapie GEO-MAP:
  * obiekty z polem A2 = DZIAŁKI (parcele),
  * obiekty z polem A1 = WYDZIELENIA (np. "3a", "12a").

Format pliku .VAL czytany przez rozliczanie (excel_tasks.wczytaj_i_przetworz_val):
  ^<litera wydzielenia> <pole m2> ...   <- wydzielenie na tej działce
  ^XXXX... <pole m2> ...                <- nagłówek sumy działki (ignorowany)
  ;----...
  *<numer działki> <pole m2> ...        <- działka
"""
import pathlib
import re

from app.core.leniwe_importy import leniwy_modul

shapely_geometry = leniwy_modul("shapely.geometry")

_KRESKA = ";" + "-" * 64


def _pisz(log, tekst):
    """Dopisz komunikat do logu.

    `log` może być LISTĄ (log.append) albo FUNKCJĄ (log(tekst)) — dzięki temu
    moduł działa i z klasycznym zbieraniem komunikatów, i z zakładką, która
    podaje własną funkcję piszącą.
    """
    if log is None:
        return
    if hasattr(log, "append"):
        log.append(tekst)
    elif callable(log):
        log(tekst)


def _czytaj(plik):
    """Zwraca listę obiektów mapy: {a1, a2, pts}."""
    t = pathlib.Path(plik).read_bytes().decode("cp1250", errors="replace")
    obiekty = []
    cur = None
    for l in t.replace("\r\n", "\n").split("\n"):
        if l.startswith("*"):
            if cur is not None:
                obiekty.append(cur)
            cur = {"pts": [], "a1": "", "a2": ""}
        if cur is None:
            continue
        if l.startswith("P "):
            q = l.split()
            try:
                cur["pts"].append((float(q[2]), float(q[3])))
            except (IndexError, ValueError):
                pass
        elif l.startswith(":A1["):
            cur["a1"] = l[4:-1].strip()
        elif l.startswith(":A2["):
            cur["a2"] = l[4:-1].strip()
    if cur is not None:
        obiekty.append(cur)
    return obiekty


def _poligon(pts):
    if len(pts) < 3:
        return None
    try:
        p = shapely_geometry.Polygon(pts)
        if not p.is_valid:
            p = p.buffer(0)
        return p if not p.is_empty else None
    except Exception:                                       # noqa: BLE001
        return None


def _czy_wydzielenie(a1):
    return bool(a1) and "X" not in a1 and re.match(r"^\d+[a-z]", a1) is not None


def _czy_dzialka(a2):
    """Czy pole A2 to naprawdę NUMER DZIAŁKI (np. 80, 24/2, 123a)?

    W mapach GEO-MAP w polu A2 bywają też nazwy i etykiety (nazwa wsi,
    „Ls", „dr", „R") — te NIE są działkami i nie mogą trafić do .VAL.
    """
    a2 = (a2 or "").strip()
    if not a2:
        return False
    return re.match(r"^\d+(/\d+)?[a-z]?$", a2, re.IGNORECASE) is not None


def generuj_val(plik_map, plik_val=None, log=None):
    """Tworzy plik .VAL z mapy .MAP. Zwraca słownik z wynikiem."""
    def _log(s):
        _pisz(log, s)

    plik_map = pathlib.Path(plik_map)
    if plik_val is None:
        plik_val = plik_map.with_suffix(".VAL")
    plik_val = pathlib.Path(plik_val)

    obiekty = _czytaj(plik_map)
    dzialki, wydzielenia = [], []
    for o in obiekty:
        p = _poligon(o["pts"])
        if p is None:
            continue
        if _czy_dzialka(o["a2"]):
            dzialki.append((o["a2"], p))
        if _czy_wydzielenie(o["a1"]):
            wydzielenia.append((o["a1"], p))
    _log("%s: działek %d, wydzieleń %d" % (plik_map.name, len(dzialki), len(wydzielenia)))
    if not dzialki or not wydzielenia:
        _log("  pominięto — brak działek (A2) albo wydzieleń (A1) w mapie.")
        return {"ok": False, "error": "brak działek lub wydzieleń", "dzialek": len(dzialki),
                "wydzielen": len(wydzielenia)}

    linie = ["Rozliczenie użytków w działkach z obliczeniem wartości",
             "Jednostka pola -> metry kwadratowe",
             "Bez uwzgledniania wpływu odwzorowania",
             ";----------------Oznaczenie --------Pole ---Cena -----Wartość",
             _KRESKA]
    pozycji = 0
    for nr, pd in dzialki:
        czesci = []
        for lit, pw in wydzielenia:
            try:
                if not pd.intersects(pw):
                    continue
                m2 = int(round(pd.intersection(pw).area))
            except Exception:                               # noqa: BLE001
                continue
            if m2 > 0:
                czesci.append((lit, m2))
        if not czesci:
            continue
        suma = int(round(pd.area))
        for lit, m2 in czesci:
            linie.append("^" + lit.rjust(27) + "%13d%8d%13d" % (m2, 0, 0))
            pozycji += 1
        linie.append("^" + "X" * 26 + "%13d%8d%13d" % (suma, 0, 0))
        linie.append(_KRESKA)
        linie.append("*" + "%26s%13d%21d%13d" % (nr, suma, 0, suma))
        linie.append(_KRESKA)
        linie.append(";")
        linie.append(_KRESKA)
    linie.append(";")
    plik_val.write_bytes(("\r\n".join(linie) + "\r\n").encode("cp1250", errors="replace"))
    _log("  zapisano: %s  (pozycji: %d)" % (plik_val, pozycji))
    return {"ok": True, "wynik": str(plik_val), "pozycji": pozycji,
            "dzialek": len(dzialki), "wydzielen": len(wydzielenia)}


def generuj_val_folder(folder, log=None):
    """Generuje .VAL dla wszystkich map .MAP w folderze (i podfolderach)."""
    folder = pathlib.Path(folder)
    pliki = [p for p in sorted(folder.rglob("*"))
             if p.is_file() and p.suffix.lower() == ".map"]
    _pisz(log, "Znaleziono map .MAP: %d" % len(pliki))
    wyniki = []
    for p in pliki:
        try:
            wyniki.append(generuj_val(p, log=log))
        except Exception as e:                              # noqa: BLE001
            _pisz(log, "  BŁĄD przy %s: %s" % (p.name, e))
            wyniki.append({"ok": False, "error": str(e)})
    return wyniki
