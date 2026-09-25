# -*- coding: utf-8 -*-
"""Nowe szablony wydruków: HTML → PDF bez Worda i bez COM.

Moduł samodzielny (bez zależności od GUI):

  * parsuje wyczyszczone pliki TXT mietka (te same, które szły do Worda),
  * renderuje estetyczne, oszczędne tuszowo strony A4 (HTML),
  * zamienia HTML na PDF przez wbudowaną przeglądarkę (Edge/Chrome) w trybie
    headless — na Windows 10/11 Edge jest zainstalowany fabrycznie,
  * buduje stronę tytułową (odpowiednik STR_TYT.docx) i konwertuje
    dokumenty .docx (opisy ogólne, skróty) na HTML → PDF bez uruchamiania
    Worda.

Marginesy pobierane są z tego samego źródła, co dotychczas
(margins_config / kreator Pełnego Automatu).
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# --------------------------------------------------------------- konstanty

# raporty drukowane poziomo (jak dotychczas)
POZIOMO = {"REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1"}

# raporty, z których znika kolumna nazwisk przy "usuń nazwiska"
USUWA_NAZWISKA = {"REJESTR1", "WSKAZ1"}

# domyślne marginesy [cm]: góra, prawo, dół, lewo — gdy brak w konfiguracji
_DOMYSLNE_MARGINESY = (1.3, 1.1, 1.5, 1.1)

# klucz w margins_dict (kreator) → typ raportu w tym module
_MAPA_MARGINESOW = {
    "REJESTR1": "REJESTR1", "OPTAX": "OPTAX", "TAB_KLW3": "TAB_KLW3",
    "WSKAZ1": "WSKAZ1", "HALIZNY": "HALIZNY", "WYK_NEG": "WYK_NEG",
    "ZEST1": "ZEST1", "WK_ZM1": "WK_ZM1",
    "OPIS": "WSK_ZB",          # WSK_ZB drukujemy z marginesami "Opisu"
    "SKROTY": "SKROTY",
}

NAGLOWKI = {"rejestru", "działki", "poddz.", "zalesiona", "zabiegu", "właściciela",
            "Współwłaściciele", "Oddział", "Powierzchnia", "mapie", "Numer",
            "gospod.", "wiek,", "ochr.", "Nazwisko", "oddz/pod", "nr dział.",
            "gat.", "Bon", "Rodzaj", "Zadania", "ZADANIA", "leśnej",
            "[ha]", "wydz.", "powierzchni", "[m3]", "m3"}

# --------------------------------------------------------------- parsowanie

def wczytaj(path):
    """Czyta plik TXT mietka: cp852, czyszczenie sekwencji PCL."""
    raw = Path(path).read_bytes().decode("cp852", errors="replace")
    raw = re.sub(r"\x1b[(&][0-9A-Za-z.]*[A-Za-z]", "", raw)
    raw = raw.replace("\x1b", "")
    return raw.replace("\x0c", "\n").replace("\r", "")


def komorki(linia):
    """Rozbija wiersz tabeli '│ a │ b │' na komórki."""
    if "│" not in linia:
        return None
    czesci = linia.strip().strip("│").split("│")
    return [c.strip() for c in czesci]


def czy_sep(linia):
    s = linia.strip()
    return bool(s) and set(s) <= set("┌┐└┘├┤┬┴┼─│-")


def czy_naglowek(k):
    """Wiersz powtórzonego nagłówka strony (podział strony w pliku mietka)."""
    return any(c in NAGLOWKI for c in k if c)


def fnum(x):
    x = x.strip()
    return x if x else ""


def _num_ok(v):
    return v == "" or re.match(r"^[\d.]+$", v)


def meta_z_pliku(path):
    """(obiekt, stan_na, okres) z nagłówków pliku TXT."""
    try:
        head = wczytaj(path)[:1200]
    except OSError:
        return "", "", ""
    m = re.search(r"Obiekt:\s*(\S+)", head) or re.search(r"dla obiektu\s+(\S+)", head)
    obiekt = m.group(1).upper() if m else ""
    m = re.search(r"Stan na:\s*(\S+)", head)
    stan = m.group(1) if m else ""
    m = re.search(r"na okres od (.*? do .*?)(?:\s{2,}|\r|\n|$)", head)
    if not m:
        # WSK_ZB: "w 10-leciu od 01-01-2027 do 31-12-2036 wg. wskazań..."
        m = re.search(r"w 10-leciu od (\S+) do (\S+)", head)
        okres = f"{m.group(1)} do {m.group(2)}" if m else ""
    else:
        okres = m.group(1).strip()
    return obiekt, stan, okres


def razem_z_optax(path):
    """'Razem' (ha) z OPTAX — na stronę tytułową."""
    m = re.search(r"Razem\s*[^0-9A-Za-z]*([\d\s]+(?:[.,]\d+)?)", wczytaj(path))
    return m.group(1).replace(" ", "").strip() if m else "[BRAK_DANYCH]"

# --------------------------------------------------------------- REJESTR1

def _wiersz_rejestru(k):
    """Pojedynczy wiersz danych rejestru (wydzielenie + wskazanie)."""
    return {"dz": k[2], "pod": k[3], "gat": k[4], "w": k[5], "bon": k[6],
            "zal": k[7], "odn": k[8], "poz": k[9], "inne": k[10],
            "razem": k[11] if len(k) > 11 else "",
            "gzal": k[12] if len(k) > 12 else "",
            "ochr": k[13] if len(k) > 13 else "",
            "rodzaj": k[14] if len(k) > 14 else "",
            "pow_z": k[15] if len(k) > 15 else "",
            "miaz_z": k[16] if len(k) > 16 else "",
            "wyk": k[17] if len(k) > 17 else ""}


def _ma_dane(k):
    """Czy wiersz niesie dane (wydzielenie lub wskazanie)? Pomija separatory '-'."""
    return any(c and not set(c) <= set("-") for c in k[2:18])

def _tylko_wykon(k):
    """Wiersz z samą dopiską w kolumnie 'Wykon.' — dopisz do poprzedniego."""
    return (len(k) > 17 and k[17]
            and not any(c and not set(c) <= set("-") for c in k[2:17]))


def parse_rejestr1(path):
    """Pozycja = grupa wierszy jednego nr rejestru (współwłaściciele razem).

    Struktura pliku mietka: wiersz właściciela (nr rej. + nazwisko z udziałem
    + ew. pierwsze wydzielenie/wskazanie), potem wiersze adresu (kolumna
    nazwiska) i dalszych zabiegów (Rodzaj/Pow), kolejni właściciele w tej
    samej pozycji, a na końcu wiersze „Razem dzialka/pozycja/obiekt".
    """
    lines = wczytaj(path).split("\n")
    pozycje, cur = [], None
    for ln in lines:
        k = komorki(ln)
        if k is None or len(k) < 11 or czy_sep(ln):
            continue
        if k[0] == "1" and k[1] == "2" and k[2] == "3":      # numeracja
            continue
        if czy_naglowek(k):
            continue
        a = k[0]
        if a and re.match(r"^\d+(/\d+)?$", a):
            if cur is None or cur["nr"] != a:
                cur = {"nr": a, "wlasciciele": [], "razem_d": "",
                       "razem_p": None, "razem_ob": False, "razem_d_nr": ""}
                pozycje.append(cur)
            if k[1]:                                         # nowy właściciel
                cur["wlasciciele"].append(
                    {"nazw": k[1], "adres": "", "wiersze": []})
            if _ma_dane(k) and cur["wlasciciele"]:
                cur["wlasciciele"][-1]["wiersze"].append(_wiersz_rejestru(k))
            elif (_tylko_wykon(k) and cur["wlasciciele"]
                  and cur["wlasciciele"][-1]["wiersze"]):
                w = cur["wlasciciele"][-1]["wiersze"][-1]
                w["wyk"] = (w["wyk"] + " " + k[17]).strip()
            continue
        if cur is None:
            continue
        if k[1] and "Razem dzialka" in k[1]:
            m = re.search(r"(-?[\d.]+)\s*ha", k[1])
            cur["razem_d"] = m.group(1) if m else ""
            m2 = re.search(r"Razem dzialka\s+(\S+)", k[1])
            cur["razem_d_nr"] = m2.group(1) if m2 else ""
        elif k[1] and "Razem pozycja" in k[1]:
            cur["razem_p"] = [k[7], k[8], k[9], k[10],
                              k[11] if len(k) > 11 else "",
                              k[12] if len(k) > 12 else ""]
        elif k[1] and "Razem obiekt" in k[1]:
            cur["razem_ob"] = True
            m = re.search(r"(-?[\d.]+)\s*ha", k[1])
            cur["razem_d"] = m.group(1) if m else ""
            cur["razem_ob_v"] = [k[7], k[8], k[9], k[10],
                                 k[11] if len(k) > 11 else "",
                                 k[12] if len(k) > 12 else ""]
        else:
            # wiersz adresu (nazwisko wypełnione, brak danych działki)
            # i/lub kolejny zabieg (Rodzaj/Pow/Miąż) tego samego właściciela
            if k[1] and cur["wlasciciele"]:
                wl = cur["wlasciciele"][-1]
                if not wl["adres"]:
                    wl["adres"] = k[1]
            if _ma_dane(k) and cur["wlasciciele"]:
                cur["wlasciciele"][-1]["wiersze"].append(_wiersz_rejestru(k))
            elif (_tylko_wykon(k) and cur["wlasciciele"]
                  and cur["wlasciciele"][-1]["wiersze"]):
                w = cur["wlasciciele"][-1]["wiersze"][-1]
                w["wyk"] = (w["wyk"] + " " + k[17]).strip()
    return pozycje

# --------------------------------------------------------------- OPTAX

def parse_optax(path):
    lines = wczytaj(path).split("\n")
    rek, cur = [], None
    for ln in lines:
        k = komorki(ln)
        if k is None or len(k) < 14 or czy_sep(ln):
            continue
        if czy_naglowek(k):
            continue
        if k[0].isdigit() and k[1].isdigit():
            continue
        if k[0]:
            if cur:
                rek.append(cur)
            cur = {"oddz": k[0], "pow": k[1], "opis": [k[2]], "el": k[3:11],
                   "wsk": k[11:14],
                   "wyk": [k[14] if len(k) > 14 else "",
                           k[15] if len(k) > 15 else "",
                           k[16] if len(k) > 16 else ""]}
        elif cur is not None and (k[2] or k[1]):
            cur["opis"].append(k[2] if k[2] else k[1])
    if cur:
        rek.append(cur)
    return rek

# --------------------------------------------------------------- TAB_KLW3

def parse_tabklw3(path):
    lines = wczytaj(path).split("\n")
    grupy, nazwa, czekam, buf_pow = [], "", False, None
    for ln in lines:
        k = komorki(ln)
        if k is None:
            continue
        if k[0] == "1" and k[1] == "2":
            czekam = True
            continue
        if not czekam:
            continue
        if k[0] and "├" in ln:                       # separator z gatunkiem
            nazwa = k[0]
            continue
        if len(k) < 4:
            continue
        if k[0] in ("", "Razem") and _num_ok(k[1]) and all(_num_ok(v) for v in k[2:]):
            if k[0] == "Razem":
                nazwa = "Razem"
            if buf_pow is None:
                buf_pow = k[2:]
            else:
                grupy.append((nazwa or "?", buf_pow, k[2:]))
                buf_pow, nazwa = None, ""
    return grupy

# --------------------------------------------------------------- WSKAZ1

def parse_wskaz1(path):
    lines = wczytaj(path).split("\n")
    wlasciciele, cur, tbl = [], None, False
    for ln in lines:
        if ln.startswith("P. "):
            m = re.match(r"^P\.\s+(.*?)\s{2,}.*Nr rej:\s*(\S+)", ln)
            if m:
                cur = {"nazw": m.group(1), "rej": m.group(2), "adres": "",
                       "wiersze": [], "razem": ["", ""]}
                wlasciciele.append(cur)
            tbl = False
            continue
        if ln.startswith("Adres") and cur is not None:
            cur["adres"] = ln.split(":", 1)[1].strip()
            continue
        k = komorki(ln)
        if k is None or czy_sep(ln):
            continue
        if k[0].startswith("Razem") and cur is not None:
            cur["razem"] = [k[1], k[2], k[5] if len(k) > 5 else ""]
            continue
        if len(k) < 8:
            continue
        if k[0].isdigit() and k[1].isdigit():
            tbl = True
            continue
        if czy_naglowek(k):
            continue
        if cur is not None:
            if k[1] or k[5]:
                cur["wiersze"].append({"dz": k[0], "pod": k[1], "las": k[2],
                                       "zal": k[3], "opis": k[4], "zad": k[5],
                                       "pow": k[6] if len(k) > 6 else "",
                                       "maks": k[7] if len(k) > 7 else ""})
    return wlasciciele

# --------------------------------------------------------------- WSK_ZB

def parse_wskzb(path):
    lines = wczytaj(path).split("\n")
    sekcje, cur, cur_pod = [], None, None
    for ln in lines:
        s = ln.rstrip()
        if not s.strip():
            continue
        m = re.match(r"^([IVX]+\.)\s+(.*)$", s.strip())
        if m:
            cur = {"nr": m.group(1), "tyt": m.group(2), "pod": [], "poz": []}
            sekcje.append(cur)
            cur_pod = None
            continue
        if cur is None:
            continue
        m = re.match(r"^\s*([A-Z])\.\s+(.*)$", s)
        if m:
            cur_pod = {"litera": m.group(1), "tyt": m.group(2), "poz": []}
            cur["pod"].append(cur_pod)
            continue
        m = re.match(r"^\s*(?:\d+\.)\s+(.*?)\s{2,}(-?[\d.]+)\s*ha"
                     r"(?:\s{2,}(\d+)\s*m3)?\s*$", s)
        if m:
            (cur_pod["poz"] if cur_pod else cur["poz"]).append(
                {"nazw": m.group(1), "ha": m.group(2), "m3": m.group(3) or ""})
            continue
        m = re.match(r"^\s*(Ogółem.*?|Razem.*?)\s{2,}(-?[\d.]+)\s*ha"
                     r"(?:\s{2,}(\d+)\s*m3)?\s*$", s)
        if m:
            (cur_pod if cur_pod else cur)["suma"] = {
                "nazw": m.group(1), "ha": m.group(2), "m3": m.group(3) or ""}
    return sekcje

# --------------------------------------------------------------- ZEST1/HALIZNY/WYK_NEG

def parse_zest1(path):
    rows = []
    for ln in wczytaj(path).split("\n"):
        if "|" not in ln:
            continue
        k = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(k) < 4 or (k[0] and set(k[0]) <= set("-")):
            continue
        if k[0] == "nr dział.":
            continue
        rows.append(k)
    return rows


def parse_halizny(path):
    rows = []
    for ln in wczytaj(path).split("\n"):
        k = komorki(ln)
        if k is None or czy_sep(ln) or len(k) < 3:
            continue
        if k[0].isdigit() or czy_naglowek(k):
            continue
        rows.append(k)
    return rows


def parse_wyk_neg(path):
    wiersze, razem, tytul = [], None, ""
    for ln in wczytaj(path).split("\n"):
        if not tytul and "Zestawienie" in ln and "neg" in ln:
            tytul = ln.strip()
            continue
        k = komorki(ln)
        if k is None or czy_sep(ln) or len(k) < 4:
            continue
        if czy_naglowek(k) or (k[0].isdigit() and k[1].isdigit()):
            continue
        if k[0].startswith("Razem"):
            razem = (k[1], k[2])
            continue
        if k[0]:
            wiersze.append({"wydz": k[0], "opis": k[1], "pow": k[2],
                            "zas": k[3], "uwagi": k[4] if len(k) > 4 else ""})
    return tytul, wiersze, razem

# --------------------------------------------------------------- HTML: styl

def _marginesy(margins, typ):
    """(góra, prawo, dół, lewo) w cm dla typu raportu."""
    if not margins:
        return _DOMYSLNE_MARGINESY
    klucz = typ
    for k, v in _MAPA_MARGINESOW.items():
        if v == typ:
            klucz = k
            break
    poz = margins.get(klucz) or margins.get(typ)
    if not poz:
        return _DOMYSLNE_MARGINESY
    try:
        t, b, l, r = [float(x) for x in poz[:4]]
    except (TypeError, ValueError, IndexError):
        return _DOMYSLNE_MARGINESY
    return (t, r, b, l)


CSS = """
  * { box-sizing: border-box; }
  body { font: 8.6pt/1.4 "Segoe UI", Arial, sans-serif; color: #141414;
         margin: 0; background: #fff; }
  .hdr { margin-bottom: 5mm; border-bottom: 1.6pt solid #222; padding-bottom: 2mm; }
  .hdr .agencja { font-size: 7.5pt; letter-spacing: .4px; color: #444; }
  .hdr h1 { font-size: 12pt; margin: 1mm 0 .5mm; letter-spacing: .2px; }
  .hdr .meta { font-size: 8pt; color: #333; }
  .hdr .meta b { color: #111; }
  table { border-collapse: collapse; width: 100%; }
  thead { display: table-header-group; }
  th { font-weight: 600; font-size: 7.6pt; background: #f2f2f2; }
  /* pionowy napis ("Ochr.") — transform na spanie, NIE na komórce:
     transform bezpośrednio na th psuł obramowanie przy border-collapse */
  th.vcol { padding: 2px 0; }
  .vtxt { display: inline-block; writing-mode: vertical-rl;
          text-orientation: mixed; transform: rotate(180deg); }

  th, td { border: 0.4pt solid #9a9a9a; padding: 2.4px 5px; vertical-align: top; }
  td.n, th.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  td.c { text-align: center; }
  tr.razem td { font-weight: 600; background: #f7f7f7; }
  tr.razem.ob td { border-top: 1.6pt solid #111; font-weight: 700; }
  tr.sub td { font-style: italic; background: #fbfbfb; }
  td.rl { text-align: right; white-space: nowrap; }
  tr.rdz td { background: none; border-top-color: transparent; border-bottom-color: transparent; }
  tr.rdz td.rl { text-align: right; font-style: italic; font-weight: 600;
                 white-space: nowrap; color: #1a1a1a; }
  .grupa td { border-top: 1.1pt solid #555; }
  .opis { color: #333; }
  .wlasc .nazw { font-weight: 600; }
  .foot { margin-top: 4mm; font-size: 7pt; color: #666;
          border-top: 0.5pt solid #aaa; padding-top: 1mm; }
  @media screen { body { max-width: 200mm; margin: 10px auto; padding: 0 12px;
                         background: #e8e8e8; } }
"""


def _as_float(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None

def _css_czcionek(cz):
    """CSS z ustawień czcionek: {"tytul": {"pt": 12, "font": "Arial"},
    "tabela": {...}}. Tytuł = nagłówek dokumentu (.hdr h1); tabela = cała
    treść tabel. Obiekt, "Stan na" i AGENCJA zostają z oryginalną czcionką.
    Wartości równe wbudowanym domyślnym (12 / 8.6 pt) nie wstrzykują stylu,
    więc niezmienione ustawienia zachowują dzisiejszy wygląd 1:1."""
    if not isinstance(cz, dict):
        return ""
    out = []
    for sekcja, domysl, sel in (("tytul", 12.0, ".hdr h1"),
                                ("tabela", 8.6, "table th, table td")):
        c = cz.get(sekcja) or {}
        if not isinstance(c, dict):
            continue
        pt = _as_float(c.get("pt"))
        fam = str(c.get("font") or "").replace('"', "").strip()
        dek = []
        if pt is not None and (abs(pt - domysl) > 1e-9 or fam):
            dek.append(f"font-size: {pt:g}pt")
        if fam:
            dek.append(f'font-family: "{fam}", Arial, sans-serif')
        if dek:
            out.append(f"{sel} {{ " + "; ".join(dek) + "; }")
    return "\n".join(out)

def _strona(tytul, obiekt, stan, tresc, extra_css="", poziom=False,
            marginesy=None, agencja="AGENCJA „CEZAR”", tytul2="",
            czcionki=None):
    t, r, b, l = marginesy or _DOMYSLNE_MARGINESY
    orient = "@page { size: A4 landscape; }" if poziom else ""
    page = f"@page {{ size: A4; margin: {t}cm {r}cm {b}cm {l}cm; }}"
    cz = _css_czcionek(czcionki)
    return f"""<!DOCTYPE html>
<html lang="pl"><head><meta charset="utf-8">
<title>{tytul} — {obiekt}</title>
<style>{page}{orient}{CSS}{extra_css}{cz}</style></head>
<body>
<div class="hdr">
  <div class="agencja">{agencja}</div>
  <h1>{tytul}{('<br>' + tytul2) if tytul2 else ''}</h1>
  <div class="meta">Obiekt: <b>{obiekt}</b>{(" &nbsp;—&nbsp; Stan na: <b>" + stan + "</b>") if stan else ""}</div>
</div>
{tresc}
</body></html>"""

# --------------------------------------------------------------- HTML: raporty

def html_rejestr1(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    pozycje = parse_rejestr1(path)
    tr = []
    for p in pozycje:
        wlasciciele = p["wlasciciele"] or [{"nazw": "", "adres": "",
                                            "wiersze": []}]
        n_rows = sum(max(len(wl["wiersze"]), 1) for wl in wlasciciele)
        nr_first = True
        for wl in wlasciciele:
            wiersze = wl["wiersze"] or [None]      # właściciel bez własnych wierszy
            wl_first = True
            for w in wiersze:
                if w is not None:
                    kom = (f'<td class="c">{w["dz"]}</td><td class="c">{w["pod"]}</td>'
                           f'<td class="c">{w["gat"]}</td><td class="n">{w["w"]}</td>'
                           f'<td class="c">{w["bon"]}</td><td class="n">{w["zal"]}</td>'
                           f'<td class="n">{w["odn"]}</td><td class="n">{w["poz"]}</td>'
                           f'<td class="n">{w["inne"]}</td><td class="n">{w["razem"]}</td>'
                           f'<td class="n">{w["gzal"]}</td><td class="n">{w["ochr"]}</td>'
                           f'<td>{w["rodzaj"]}</td><td class="n">{w["pow_z"]}</td>'
                           f'<td class="n">{w["miaz_z"]}</td><td>{w["wyk"]}</td>')
                else:
                    kom = "<td></td>" * 16
                wl_td = ""
                if wl_first and not bez_nazwisk:
                    tresc_wl = ""
                    if wl["nazw"]:
                        tresc_wl += f'<div class="nazw">{wl["nazw"]}</div>'
                    if wl["adres"]:
                        tresc_wl += f'<div class="opis adr">{wl["adres"]}</div>'
                    wl_td = (f'<td rowspan="{max(len(wl["wiersze"]), 1)}" '
                             f'class="wlasc">{tresc_wl}</td>')
                if nr_first:
                    tr.append('<tr class="grupa">'
                              f'<td rowspan="{n_rows}">{p["nr"]}</td>'
                              + wl_td + kom + "</tr>")
                    nr_first = False
                else:
                    tr.append("<tr>" + wl_td + kom + "</tr>")
                wl_first = False
        cd = 15 if bez_nazwisk else 16
        if p["razem_d"] and not p["razem_ob"]:
            nr_d = p.get("razem_d_nr") or ""
            tr.append(f'<tr class="rdz">'
                      f'<td class="rl" colspan="2">Razem działka {nr_d} — {p["razem_d"]} ha</td>'
                      f'<td colspan="{cd}"></td></tr>')
        cp_ = 4 if bez_nazwisk else 5
        if p["razem_p"]:
            rp = list(p["razem_p"]) + [""] * (6 - len(p["razem_p"]))
            tr.append(f'<tr class="sub"><td class="rl" colspan="2">Razem pozycja</td>'
                      f'<td colspan="{cp_}"></td>'
                      f'<td class="n">{rp[0]}</td><td class="n">{rp[1]}</td>'
                      f'<td class="n">{rp[2]}</td><td class="n">{rp[3]}</td>'
                      f'<td class="n">{rp[4]}</td><td class="n">{rp[5]}</td>'
                      f'<td colspan="5"></td></tr>')
        if p["razem_ob"]:
            rv = list(p.get("razem_ob_v") or []) + [""] * 6
            tr.append(f'<tr class="razem ob">'
                      f'<td class="rl" colspan="2">Razem obiekt — {p["razem_d"]} ha</td>'
                      f'<td colspan="{cp_}"></td>'
                      f'<td class="n">{rv[0]}</td><td class="n">{rv[1]}</td>'
                      f'<td class="n">{rv[2]}</td><td class="n">{rv[3]}</td>'
                      f'<td class="n">{rv[4]}</td><td class="n">{rv[5]}</td>'
                      f'<td colspan="5"></td></tr>')
    nazw_th = "" if bez_nazwisk else (
        '<th rowspan="3" style="width:14%">Nazwisko i imię<br>'
        '<span class="opis">adres, współwłaściciele</span></th>\n')
    head = f"""<thead><tr>
<th rowspan="3" style="width:5%">Nr<br>rej.</th>
{nazw_th}<th rowspan="3" style="width:5%">Nr<br>działki</th>
<th rowspan="3" style="width:4.5%">Oddz.<br>poddz.</th>
<th colspan="8">Opis i powierzchnia lasów [ha]</th>
<th rowspan="3" class="wz">Pow.<br>gruntów<br>do zal.</th>
<th rowspan="3" class="vcol"><span class="vtxt">Ochr.</span></th>
<th colspan="3">Wskazania gospodarcze</th>
<th rowspan="3" style="width:5.5%">Wykon.</th></tr>
<tr><th colspan="4">zalesiona</th><th colspan="2">nie zalesiona</th>
<th rowspan="2">inne<br>grunty</th><th rowspan="2">razem<br>lasy</th>
<th rowspan="2" style="width:10%">Rodzaj<br>zabiegu</th>
<th rowspan="2" class="n" style="width:5.5%">Pow.<br>[ha]</th>
<th rowspan="2" class="n" style="width:5.5%">Miąż.<br>[m³]</th></tr>
<tr><th>gat.</th><th>W</th><th>Bon</th><th>Pow.</th><th>do<br>odn.</th>
<th>pozost.</th></tr></thead>"""
    tresc = ('<table>' + head + '<tbody>' + "".join(tr) + "</tbody></table>")
    return _strona("Rejestr działek leśnych i gruntów do zalesienia wg. właścicieli",
                   obiekt, stan, tresc, poziom=True, marginesy=marginesy,
                   czcionki=czcionki)


def html_optax(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    rek = parse_optax(path)
    tr = []
    for r in rek:
        opis = "<br>".join(r["opis"])
        el = r["el"]
        wsk = r["wsk"]
        wyk = r["wyk"]
        tr.append("<tr>"
                  f'<td class="c" style="font-weight:600">{r["oddz"]}</td>'
                  f'<td class="n">{r["pow"]}</td>'
                  f'<td class="opis">{opis}</td>'
                  f'<td class="c">{el[0]}</td><td class="n">{el[1]}</td>'
                  f'<td class="c">{el[2]}</td><td class="n">{el[3]}</td>'
                  f'<td class="n">{el[4]}</td><td class="c">{el[5]}</td>'
                  f'<td class="n">{el[6]}</td><td class="n">{el[7]}</td>'
                  f'<td>{wsk[0]}</td><td class="n">{wsk[1]}</td>'
                  f'<td class="n">{wsk[2]}</td>'
                  f'<td>{wyk[0]}</td><td class="n">{wyk[1]}</td>'
                  f'<td class="n">{wyk[2]}</td></tr>')
    tresc = """<table>
<thead><tr>
<th rowspan="2" style="width:5%">Oddział<br>poddz.</th>
<th rowspan="2" class="n" style="width:6%">Pow.<br>[ha]</th>
<th rowspan="2" style="width:30%">Opis taksacyjny lasu, gruntu<br>
przeznaczonego do zalesienia</th>
<th colspan="8">Elementy taksacyjne</th>
<th colspan="3">Wskazania gospodarcze</th>
<th colspan="3">Wykonanie</th></tr>
<tr>
<th style="width:5%">Gat. gł.</th><th class="n" style="width:4%">Wiek</th>
<th style="width:5%">Klasa wieku</th><th class="n" style="width:4%">Wys. [m]</th>
<th class="n" style="width:4%">Pierw. [cm]</th><th class="n" style="width:4%">Bon.</th>
<th>Zad.</th><th class="n">Miąż.<br>na pow.<br>[m3]</th>
<th style="width:11%">Rodzaj wskazania</th><th class="n" style="width:6%">Pow.<br>[ha]</th>
<th class="n" style="width:5%">Maks.<br>miąż. do<br>pozysk.<br>[m3]</th>
<th>Czynn.</th><th class="n" style="width:4%">[ha]</th><th class="n" style="width:4%">[m3]</th>

<tbody>""" + "".join(tr) + "</tbody></table>"
    return _strona("Opis lasów i gruntów przeznaczonych do zalesienia",
                   obiekt, stan, tresc, poziom=True, marginesy=marginesy,
                   czcionki=czcionki)


def html_tabklw3(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    grupy = parse_tabklw3(path)
    KLASY = ["Ia", "Ib", "IIa", "IIb", "IIIa", "IIIb", "IVa", "IVb",
             "Va", "Vb", "VIa", "VIb", "VII+", "K.D.O.", "K.O.", "Razem", "Ogółem"]
    tr = []
    for nazwa, pw, mz in grupy:
        razem = nazwa in ("Razem", "OGÓŁEM", "Ogółem")
        cls = "razem" if razem else "grupa"
        bold = ' style="font-weight:700"' if razem else ""
        tr.append(f'<tr class="{cls}"><td rowspan="2" class="c"{bold}>{nazwa}</td>'
                  '<td class="c opis">ha</td>' + "".join(
                      f'<td class="n">{fnum(v)}</td>' for v in pw) + "</tr>")
        tr.append('<tr><td class="c opis">m³</td>' + "".join(
            f'<td class="n">{fnum(v)}</td>' for v in mz) + "</tr>")
    head1 = ('<tr><th rowspan="3" style="width:7%">Gatunek<br>panujący</th>'
             '<th rowspan="3" class="n" style="width:5%">Przesłania<br>i nasiona</th>'
             '<th colspan="17">Powierzchnia w ha / miąższość w m³ — klasy i podklasy wieku</th></tr>')
    head2 = '<tr>' + "".join(f'<th class="n">{k}</th>' for k in KLASY) + "</tr>"
    tresc = ("<table><thead>" + head1 + head2 + "</thead><tbody>" + "".join(tr)
             + "</tbody></table>")
    return _strona("Zestawienie powierzchni gruntów i miąższości drzewostanu "
                   "wg gatunków panujących (głównych) wg klas i podklas wieku",
                   obiekt, stan, tresc, poziom=True, marginesy=marginesy,
                   czcionki=czcionki)


def html_wskaz1(path, obiekt, stan, okres="", bez_nazwisk=False, marginesy=None, czcionki=None):
    dane = parse_wskaz1(path)
    czesci = []
    for w in dane:
        tr = []
        for rz in w["wiersze"]:
            tr.append(f'<tr class="grupa"><td>{rz["dz"]}</td><td class="c">{rz["pod"]}</td>'
                      f'<td class="n">{rz["las"]}</td><td class="n">{rz["zal"]}</td>'
                      f'<td class="c">{rz["opis"]}</td><td>{rz["zad"]}</td>'
                      f'<td class="n">{rz["pow"]}</td><td class="n">{rz["maks"]}</td></tr>')
        tr.append(f'<tr class="razem"><td colspan="2">Razem:</td>'
                  f'<td class="n">{w["razem"][0]}</td><td class="n">{w["razem"][1]}</td>'
                  f'<td colspan="3"></td><td class="n">'
                  f'{w["razem"][2] if len(w["razem"]) > 2 else ""}</td></tr>')
        if bez_nazwisk:
            nagl = (f'<div class="wl-nag"><span class="rej">Nr rej.: '
                    f'<b>{w["rej"]}</b></span></div>')
        else:
            nagl = (f'<div class="wl-nag"><span class="nazw">P. {w["nazw"]}</span>'
                    f'<span class="adres">Adres: {w["adres"]}</span>'
                    f'<span class="rej">Nr rej.: <b>{w["rej"]}</b></span></div>')
        if not w["wiersze"]:
            czesci.append(f'<div class="wlasiciel pusty">{nagl}</div>')
            continue
        czesci.append(f'<div class="wlasiciel">{nagl}'
                      '<table><thead><tr>'
                      '<th style="width:7%">Nr<br>działki</th><th style="width:6%">Oddz.<br>poddz.</th>'
                      '<th class="n" style="width:6%">Pow. lasu<br>[ha]</th><th class="n" style="width:6%">Gruntu<br>do zal. [ha]</th>'
                      '<th style="width:13%">Opis lasu<br><span class="opis">gat. gł., wiek, bon., pow. ochr.</span></th>'
                      '<th style="width:37%">Zadania w zakresie gospodarki leśnej — rodzaj</th>'
                      '<th class="n" style="width:7%">Pow. w [ha]</th><th class="n" style="width:7%">Maks. miąż. [m³]</th>'
                      '</tr></thead><tbody>' + "".join(tr) + "</tbody></table></div>")
    extra = """
  .wlasiciel { margin-bottom: 7mm; page-break-inside: avoid; }
  .wl-nag { display: flex; gap: 5mm; font-size: 8.6pt; margin: 1.5mm 0 1mm;
            border-bottom: 0.8pt solid #666; padding-bottom: 0.8mm; }
  .wl-nag .nazw { font-weight: 700; }
  .wl-nag .rej { margin-left: auto; }
  .wlasiciel.pusty { margin-bottom: 2.5mm; }
  h2.tyt { text-align: center; font-size: 11pt; margin: 2mm 0 4mm; }
  h2.tyt span { font-size: 9pt; font-weight: 500; }"""
    okres_t = f"<br><span>na okres od {okres}</span>" if okres else ""
    tresc = ('<h2 class="tyt">ZADANIA W ZAKRESIE GOSPODARKI LEŚNEJ' + okres_t
             + "</h2>" + "".join(czesci))
    return _strona("Zadania w zakresie gospodarki leśnej",
                   obiekt, "", tresc, extra_css=extra, poziom=True,
                   marginesy=marginesy, czcionki=czcionki)


def html_wskzb(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    sekcje = parse_wskzb(path)
    czesci = []
    for s in sekcje:
        inner = ""
        def pozycje(poz, suma=None):
            h = ""
            for p in poz:
                h += (f'<tr><td>{p["nazw"]}</td><td class="n">{p["ha"]}</td>'
                      f'<td class="n">{p["m3"]}</td></tr>')
            if suma:
                h += (f'<tr class="razem"><td>{suma["nazw"]}</td>'
                      f'<td class="n">{suma["ha"]}</td><td class="n">{suma["m3"]}</td></tr>')
            return h
        if s["poz"]:
            inner += ('<table class="gl"><thead><tr><th>Pozycja</th>'
                      '<th class="n">Powierzchnia [ha]</th><th class="n">Miąższość [m3]</th>'
                      '</tr></thead><tbody>'
                      + pozycje(s["poz"], s.get("suma")) + "</tbody></table>")
        for p in s["pod"]:
            inner += (f'<div class="podt">{p["litera"]}. {p["tyt"]}</div>'
                      '<table class="gl"><tbody>' + pozycje(p["poz"], p.get("suma"))
                      + "</tbody></table>")
        czesci.append(f'<div class="sekcja"><h2>{s["nr"]} {s["tyt"]}</h2>{inner}</div>')
    tresc = "".join(czesci)
    extra = """
  .podtyt { text-align: center; font-size: 9.5pt; margin: 0 0 5mm; }
  .sekcja h2 { font-size: 10.5pt; border-bottom: 1pt solid #333;
               padding-bottom: 1mm; margin: 5mm 0 2mm; }
  .podt { font-weight: 600; margin: 3mm 0 1mm; }
  table.gl { width: 75%; margin-left: 8mm; }
  table.gl td:first-child { width: 60%; }"""
    tytul2 = (f"w 10-leciu od {stan} wg. wskazań gospodarczych" if stan else "")
    return _strona("Zestawienie czynności gospodarczych projektowanych do wykonania",
                   obiekt, "", tresc, extra_css=extra, marginesy=marginesy,
                   tytul2=tytul2, czcionki=czcionki)


def html_zest1(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    tr = []
    for k in parse_zest1(path):
        tr.append(f'<tr><td>{k[0]}</td><td class="n">{k[1]}</td>'
                  f'<td class="c">{k[2]}</td><td class="n">{k[3]}</td></tr>')
    tresc = ('<table><thead><tr><th style="width:12%">Nr działki</th>'
             '<th style="width:15%">Nr rejestru</th>'
             '<th style="width:12%">Oddz./poddz.</th>'
             '<th class="n" style="width:14%">Powierzchnia działki [ha]</th>'
             '</tr></thead><tbody>' + "".join(tr) + "</tbody></table>")
    return _strona("Skorowidz działek", obiekt, stan, tresc, marginesy=marginesy, czcionki=czcionki)


def html_halizny(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    tr = []
    for k in parse_halizny(path):
        if k[0].startswith("R.oddz"):
            tr.append(f'<tr class="sub"><td><b>R. oddz.</b></td>'
                      f'<td class="n"><b>{k[1]}</b></td>'
                      f'<td>{k[2]}</td></tr>')
        elif k[0].startswith("Razem"):
            tr.append(f'<tr class="razem"><td>Razem</td><td class="n">{k[1]}</td>'
                      f'<td>{k[2]}</td></tr>')
        elif k[0]:
            tr.append(f'<tr><td class="c" style="font-weight:600">{k[0]}</td>'
                      f'<td class="n">{k[1]}</td><td>{k[2]}</td></tr>')
    tresc = ('<table><thead><tr><th style="width:14%">Oddział poddz.</th>'
             '<th class="n" style="width:14%">Pow. wydz. [ha]</th>'
             '<th>Rodzaj powierzchni</th></tr></thead><tbody>'
             + "".join(tr) + "</tbody></table>")
    return _strona("Zestawienie powierzchni leśnych niezalesionych",
                   obiekt, stan, tresc, marginesy=marginesy, czcionki=czcionki)


def html_wyk_neg(path, obiekt, stan, bez_nazwisk=False, marginesy=None, czcionki=None):
    tytul, wiersze, razem = parse_wyk_neg(path)
    m = re.search(r"-\s*(\S+)\s*$", tytul)
    obreb = m.group(1) if m else obiekt
    tr = []
    for w in wiersze:
        tr.append(f'<tr><td class="c" style="font-weight:600">{w["wydz"]}</td>'
                  f'<td class="c">{w["opis"]}</td><td class="n">{w["pow"]}</td>'
                  f'<td class="n">{w["zas"]}</td><td>{w["uwagi"]}</td></tr>')
    if razem:
        tr.append(f'<tr class="razem"><td colspan="2">Razem</td>'
                  f'<td class="n">{razem[0]}</td><td class="n">{razem[1]}</td><td></td></tr>')
    tresc = ('<table><thead><tr><th style="width:9%">Oddz.<br>Podod.</th>'
             '<th style="width:24%">Skrócony opis lasu<br>'
             '<span class="opis">gat. pan. – bon. – wiek</span></th>'
             '<th class="n" style="width:12%">Powierzchnia [ha]</th>'
             '<th class="n" style="width:12%">Zasobność [m³]</th>'
             '<th>Uwagi</th></tr></thead><tbody>'
             + "".join(tr) + "</tbody></table>")
    return _strona("Zestawienie powierzchni i zasobności dla drzewostanów "
                   "negatywnych i źle produkujących",
                   obreb, "", tresc, marginesy=marginesy, czcionki=czcionki)


# typ -> renderer; wszystkie mają sygnaturę (path, obiekt, stan, ...)
RENDERERY = {
    "REJESTR1": html_rejestr1,
    "OPTAX": html_optax,
    "TAB_KLW3": html_tabklw3,
    "WSKAZ1": html_wskaz1,
    "WSK_ZB": html_wskzb,
    "ZEST1": html_zest1,
    "HALIZNY": html_halizny,
    "WYK_NEG": html_wyk_neg,
}

# --------------------------------------------------------------- HTML: skróty
def _skroty_sekcje(docx_path):
    """Sekcje z pliku 'Skróty i symbole': [(tytuł, [(skrót, znaczenie)...])].

    Plik ma prostą budowę: akapit z tytułem sekcji, potem tabela
    1-wierszowa o 2 kolumnach (po lewej skróty, po prawej rozwinięcia,
    każdy wiersz to jedna para).
    """
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    doc = Document(str(docx_path))
    sekcje, aktualny = [], "Skróty i symbole"
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            t = Paragraph(child, doc).text.strip()
            if t:
                aktualny = t
        elif child.tag.endswith("}tbl"):
            tab = Table(child, doc)
            for row in tab.rows:
                kom = [c.text for c in row.cells]
                if len(kom) < 2:
                    continue
                lewe = [x.strip() for x in kom[0].split("\n") if x.strip()]
                prawe = [x.strip() for x in kom[1].split("\n") if x.strip()]
                pary = list(zip(lewe, prawe))
                if pary:
                    sekcje.append((aktualny, pary))
    return sekcje

def html_skroty(docx_path, obiekt="", stan="", bez_nazwisk=False,
                marginesy=None, czcionki=None):
    """'Skróty i symbole' nowym wyglądem — HTML → PDF, bez uruchamiania Worda.

    Czyta sekcje wprost z pliku .docx (domyślnego albo własnego użytkownika),
    więc własne wersje też dostają nowy wygląd. Gdy struktura pliku jest
    inna (brak rozpoznanych sekcji) — zgłaszamy błąd i wywołujący ma wrócić
    do konwersji przez Worda.
    """
    import html as _html

    def _e(x):
        return _html.escape(str(x), quote=False)

    sekcje = _skroty_sekcje(docx_path)
    if not sekcje:
        raise ValueError(f"Nie rozpoznano sekcji skrótów w {docx_path}")
    czesci = []
    for tytul, pary in sekcje:
        tr = "".join(
            f'<tr><td class="sk">{_e(skr)}</td><td>{_e(zn)}</td></tr>'
            for skr, zn in pary)
        czesci.append(f'<h2>{_e(tytul)}</h2>'
                      f'<table class="skroty"><tbody>{tr}</tbody></table>')
    extra_css = """
  h2 { font-size: 9.6pt; text-transform: uppercase; letter-spacing: .8px;
       color: #1f3d2b; margin: 5.2mm 0 1.6mm; padding-bottom: .9mm;
       border-bottom: 1pt solid #1f3d2b; font-weight: 700; }
  h2:first-child { margin-top: 0; }
  table.skroty { border-collapse: collapse; width: 100%; margin: 0 0 2mm; }
  table.skroty td { border: 0; padding: 1.05mm 3mm; font-size: 8.4pt;
                    border-bottom: .3pt solid #d8d8d8; }
  table.skroty tr:nth-child(even) td { background: #f4f6f4; }
  table.skroty td.sk { font-weight: 600; text-align: center; min-width: 12mm;
                       white-space: nowrap; }
  table.skroty tr:last-child td { border-bottom: .5pt solid #999; }
"""
    return _strona("Wykaz skrótów i symboli", obiekt, stan, "".join(czesci),
                   extra_css=extra_css, marginesy=marginesy, czcionki=czcionki)

# --------------------------------------------------------------- HTML: strona tytułowa

WYKONAWCA = ["WYKONAWCA", "pracownia urządzania lasu", "ul. Boczna 28",
             "05-300 Mińsk Mazowiecki", "tel.25 - 759 04 94",
             "e-mail: agencja.cezar@interia.pl"]

STR_TYT_CSS = """
  * { box-sizing: border-box; }
  body { font-family: "Times New Roman", serif; color: #000; margin: 0; }
  .tstr { display: flex; flex-direction: column; height: 257mm; }
  .tyt { text-align: center; margin-top: 42mm; }
  .tyt h1 { font-size: 17pt; font-weight: 700; margin: 0 0 8mm;
            letter-spacing: .5px; }
  .tyt p { font-size: 12.5pt; margin: 0 0 4.5mm; }
  .tyt p.wieś { font-weight: 700; font-size: 14.5pt; margin: 7mm 0 7mm; }
  .tyt p.maly { font-size: 12pt; }
  .wyk { margin-top: auto; text-align: center; font-size: 11.5pt;
         line-height: 1.55; }
  .wyk b { letter-spacing: .4px; }
"""


def generuj_str_tyt_html(out_path, doc_type="UPUL", prefix="",
                         woj="", powiat="", gmina="", stan_na="", okres="",
                         village="NAZWA WSI", agencja="AGENCJA „CEZAR”"):
    """Strona tytułowa — odpowiednik STR_TYT.docx (bez wiersza powierzchni)."""
    if doc_type == "ISL":
        linie_tyt = ["INWENTARYZACJA STANU LASU",
                     "dla lasów niestanowiących własności Skarbu Państwa"]
    else:
        linie_tyt = ["UPROSZCZONY PLAN URZĄDZANIA LASÓW",
                     "nie stanowiących własności Skarbu Państwa"]
    if prefix.strip().rstrip(":").lower() == "obręb":
        linie_tyt += ["Obręb:"]
    else:
        linie_tyt += ["położonych na terenie", "obrębu"]
    wyś = (f'<div class="tstr"><div class="tyt">'
           f'<h1>{linie_tyt[0]}</h1>'
           + "".join(f'<p>{l}</p>' for l in linie_tyt[1:])
           + f'<p class="wieś">{village}</p>'
           + (f'<p>gmina {gmina}</p>' if gmina else "")
           + (f'<p>POWIAT {powiat}</p>' if powiat else "")
           + (f'<p>WOJ. {woj}</p>' if woj else "")
           + (f'<p class="maly">wg stanu na {stan_na}</p>' if stan_na else "")
           + (f'<p class="maly">na okres {okres}</p>' if okres else "")
           + '</div><div class="wyk">'
           + "<br>".join(f"<b>{WYKONAWCA[0]}</b>" if i == 0 else l
                        for i, l in enumerate(WYKONAWCA))
           + "</div></div>")
    html = f"""<!DOCTYPE html>
<html lang="pl"><head><meta charset="utf-8">
<title>Strona tytułowa — {village}</title>
<style>@page {{ size: A4; margin: 2.5cm 2cm 2cm 2cm; }}
{STR_TYT_CSS}</style></head>
<body>{wyś}</body></html>"""
    Path(out_path).write_text(html, encoding="utf-8")
    return out_path

# --------------------------------------------------------------- DOCX → HTML

_DOCX_CSS = """
  body { font: 10.5pt/1.45 "Times New Roman", serif; color: #111; }
  p { margin: 0 0 2.2mm; }
  p.c { text-align: center; } p.r { text-align: right; } p.j { text-align: justify; }
  h1,h2,h3 { margin: 3mm 0 2mm; }
  table { border-collapse: collapse; width: 100%; margin: 2mm 0; }
  td, th { border: 0.4pt solid #999; padding: 1.6px 5px; vertical-align: top; }
"""


def _runy_na_html(paragraph):
    out = []
    for r in paragraph.runs:
        t = (r.text or "")
        if not t:
            continue
        if r.bold:
            t = f"<b>{t}</b>"
        if r.italic:
            t = f"<i>{t}</i>"
        if r.underline:
            t = f"<u>{t}</u>"
        out.append(t)
    return "".join(out) if out else ""


def docx_na_html(docx_path, out_path=None, tytul=None):
    """Konwertuje .docx na prosto, wiernie wyglądający HTML (bez Worda)."""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(docx_path))
    czesci = []
    body = doc.element.body
    for el in body.iterchildren():
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = Paragraph(el, doc)
            txt = _runy_na_html(p) or ""
            styl = (p.style.name if p.style is not None else "") or ""
            if not txt.strip() and not txt:
                czesci.append("")
                continue
            if styl.startswith("Heading") or styl.startswith("Nagłówek"):
                czesci.append(f"<h3>{txt}</h3>")
                continue
            align = {1: "c", 2: "r", 3: "j"}.get(p.alignment)
            cls = f' class="{align}"' if align else ""
            czesci.append(f"<p{cls}>{txt}</p>")
        elif tag == "tbl":
            t = Table(el, doc)
            rows = []
            for row in t.rows:
                cells = []
                for c in row.cells:
                    inner = "<br>".join(
                        (_runy_na_html(p) or "&nbsp;") for p in c.paragraphs)
                    cells.append(f"<td>{inner}</td>")
                rows.append("<tr>" + "".join(cells) + "</tr>")
            czesci.append('<table>' + "".join(rows) + "</table>")
    tytul = tytul or Path(docx_path).stem
    html = (f'<!DOCTYPE html><html lang="pl"><head><meta charset="utf-8">'
            f'<title>{tytul}</title>'
            f'<style>@page {{ size: A4; margin: 2cm 1.8cm 2cm 1.8cm; }}'
            f'{_DOCX_CSS}</style></head><body>'
            + "".join(czesci) + "</body></html>")
    if out_path:
        Path(out_path).write_text(html, encoding="utf-8")
        return out_path
    return html

# --------------------------------------------------------------- HTML → PDF

_WINDOWS_PRZEGLADARKI = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
_POSIX_PRZEGLADARKI = ["chromium", "chromium-browser", "google-chrome",
                       "msedge", "chrome"]


def znajdz_przegladarke():
    for p in _WINDOWS_PRZEGLADARKI:
        if os.path.isfile(p):
            return p
    for nazwa in _POSIX_PRZEGLADARKI:
        w = shutil.which(nazwa)
        if w:
            return w
    return None


def html_na_pdf(html_path, pdf_path, timeout=120):
    """HTML → PDF przez Edge/Chrome headless. Zwraca True/False."""
    html_path, pdf_path = Path(html_path), Path(pdf_path)
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            pass
    exe = znajdz_przegladarke()
    if not exe:
        raise RuntimeError("Nie znaleziono przeglądarki (Edge/Chrome) "
                           "do wydruku PDF.")
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="forestly_pdf_") as prof:
        cmd = [exe, "--headless", "--disable-gpu", "--no-first-run",
               "--no-pdf-header-footer", "--print-to-pdf-no-header",
               f"--user-data-dir={prof}",
               f"--print-to-pdf={pdf_path.resolve()}",
               html_path.resolve().as_uri()]
        try:
            subprocess.run(cmd, capture_output=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            pass
    ok = pdf_path.exists() and pdf_path.stat().st_size > 100
    if not ok:
        raise RuntimeError(f"Nie udało się wygenerować PDF: {pdf_path.name}")
    return True


def generuj_raport_pdf(typ, txt_path, pdf_path, bez_nazwisk=False,
                       margins=None, agencja="AGENCJA „CEZAR”",
                       czcionki=None):
    """TXT mietka → HTML → PDF dla jednego raportu.

    'czcionki' = {TYP: {"tytul": {"pt":…, "font":…}, "tabela": {…}}}
    (ustawienia z kreatora / zakładki Nowe Szablony)."""
    typ = typ.upper()
    renderer = RENDERERY.get(typ)
    if renderer is None:
        raise ValueError(f"Nieznany typ raportu: {typ}")
    obiekt, stan, okres = meta_z_pliku(txt_path)
    mg = _marginesy(margins, typ)
    cz = (czcionki or {}).get(typ) if isinstance(czcionki, dict) else None
    if typ == "WSKAZ1":
        html = renderer(txt_path, obiekt, stan, okres=okres,
                        bez_nazwisk=bez_nazwisk, marginesy=mg, czcionki=cz)
    elif typ == "WSK_ZB":
        html = renderer(txt_path, obiekt, okres or stan,
                        bez_nazwisk=bez_nazwisk, marginesy=mg, czcionki=cz)
    else:
        html = renderer(txt_path, obiekt, stan,
                        bez_nazwisk=bez_nazwisk, marginesy=mg, czcionki=cz)
    with tempfile.TemporaryDirectory(prefix="forestly_tpl_") as tmp:
        html_path = Path(tmp) / f"{typ}.html"
        html_path.write_text(html, encoding="utf-8")
        html_na_pdf(html_path, pdf_path)
    return pdf_path


def docx_na_pdf(docx_path, pdf_path, margins=None):
    """DOCX → HTML → PDF bez Worda (opisy ogólne, skróty)."""
    with tempfile.TemporaryDirectory(prefix="forestly_d2p_") as tmp:
        html_path = Path(tmp) / (Path(docx_path).stem + ".html")
        docx_na_html(docx_path, html_path)
        html_na_pdf(html_path, pdf_path)
    return pdf_path
