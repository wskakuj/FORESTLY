# -*- coding: utf-8 -*-
"""
Forestly — spis treści scalonego PDF.

Strona ze spisem dokumentów wchodzących w skład scalonego pliku PDF.
Tytuły pozycji są czytane z PIERWSZYCH STRON dokumentów (a nie z nazw
plików) — dzięki temu spis pokazuje to, co widzi czytelnik, a nie to,
jak plik nazywa się na dysku.
"""

import os
import re
import tempfile
from pathlib import Path


def _czcionka():
    """Ścieżka do czcionki z polskimi znakami (None -> wbudowana helv)."""
    kandydaci = []
    try:
        from app.core.word_worker import get_resource_path
        kandydaci.append(get_resource_path("DejaVuSans.ttf"))
    except Exception:
        pass
    if os.name == "nt":
        kandydaci.append(Path(r"C:\Windows\Fonts\arial.ttf"))
    try:  # środowisko deweloperskie z matplotlibem
        import matplotlib
        kandydaci.append(Path(matplotlib.get_data_path())
                          / "fonts" / "ttf" / "DejaVuSans.ttf")
    except Exception:
        pass
    for k in kandydaci:
        try:
            if k and Path(k).exists():
                return str(k)
        except Exception:
            pass
    return None


def wyciagnij_tytul(sciezka, max_znakow=80):
    """Tytuł dokumentu = pierwsze niepuste linie tekstu z pierwszej strony."""
    try:
        try:
            import pymupdf as fitz  # PyMuPDF (nowe API)
        except ImportError:
            import fitz  # starsze wersje
        with fitz.open(str(sciezka)) as d:
            if d.page_count == 0:
                return None
            tekst = d[0].get_text("text")
    except Exception:
        return None
    linie = [re.sub(r"\s+", " ", l).strip() for l in tekst.splitlines()]
    linie = [l for l in linie if len(l) > 2]
    if not linie:
        return None
    tytul = linie[0]
    for dodatkowa in linie[1:3]:
        if len(tytul) + len(dodatkowa) + 1 > max_znakow:
            break
        tytul += " " + dodatkowa
    return tytul[:max_znakow]


def _przyjazna_nazwa(sciezka):
    """Etykieta szablonu dla pliku (fallback, gdy strona 1 jest pusta)."""
    try:
        from app.config import PDF_ORDER_TEMPLATES, template_matches
        for tpl in PDF_ORDER_TEMPLATES:
            if template_matches(tpl, Path(sciezka).name):
                return tpl["label"]
    except Exception:
        pass
    return None


def zbuduj_spis_dla(pdfs, naglowek="SPIS TREŚCI"):
    """Buduje jednostronicowy PDF ze spisem dokumentów.

    Zwraca ścieżkę do pliku tymczasowego albo None (gdy nie da się zbudować).
    Numry stron uwzględniają samą stronę spisu jako pierwszą.
    """
    try:
        try:
            import pymupdf as fitz  # PyMuPDF (nowe API)
        except ImportError:
            import fitz  # starsze wersje
    except Exception:
        return None
    try:
        from pypdf import PdfReader
    except Exception:
        return None

    pdfs = list(pdfs)
    if len(pdfs) < 2:
        return None

    wpisy = []
    offset = 1  # strona spisu
    for p in pdfs:
        p = Path(p)
        try:
            n_stron = len(PdfReader(str(p)).pages)
        except Exception:
            n_stron = 1
        tytul = wyciagnij_tytul(p) or _przyjazna_nazwa(p) or p.stem
        wpisy.append((tytul, offset))
        offset += n_stron

    czcionka = _czcionka()
    kw = {"fontname": "Fczc", "fontfile": czcionka} if czcionka else {"fontname": "helv"}

    try:
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_textbox(fitz.Rect(40, 70, 555, 115), naglowek,
                            fontsize=16, align=fitz.TEXT_ALIGN_CENTER, **kw)
        y = 150
        for i, (tytul, str_start) in enumerate(wpisy, start=1):
            pozycja = f"{i}. {tytul}  \u2014  str. {str_start}"
            rc = page.insert_textbox(fitz.Rect(50, y, 545, y + 22),
                                     pozycja, fontsize=10.5, **kw)
            y += 22 if rc >= 0 else 30  # gdyby się nie zmieściło, więcej miejsca
        fd, tmp = tempfile.mkstemp(suffix=".pdf", prefix="forestly_spis_")
        os.close(fd)
        doc.save(tmp)
        doc.close()
        return Path(tmp)
    except Exception:
        return None
