"""Słownik zamian opisów mapy — „zapamiętać?" z tabeli braków.

Użytkownik w tabeli „Sprawdź braki" może poprawić wartość, którą daje reguła
(kolumna „Co da reguła"), zaznaczyć „zapamiętać?" — a program zapamiętuje tę
zamianę i od tej pory stosuje ją sam: przy wpisywaniu opisów do mapy, przy
sprawdzaniu braków i przy sprawdzaniu zaczytania.

Kluczem zamiany jest część opisowa wartości (to, co przed ostatnim „|"), a część
z powierzchnią zostaje bez zmian. Dzięki temu jedna zamiana działa dla wszystkich
powierzchni, np.:

    "Rola i droga polna|0.20"  ->  "Rola i dr.|0.20"
    "Rola i droga polna|0.55"  ->  "Rola i dr.|0.55"

Zamiany trzymane są w pliku ``config/zamiany_opisow.json`` obok aplikacji.

Zależności: tylko biblioteka standardowa (json, pathlib) — moduł nie importuje
niczego z GUI, więc można go wołać także z rdzenia.
"""

import json
import sys
from pathlib import Path

__all__ = ["wczytaj", "zapisz", "zastosuj", "zapamietaj", "usun", "wyczysc",
           "sciezka", "klucz", "OPIS"]


def _app_dir():
    """Katalog aplikacji (obok exe po spakowaniu, główny katalog w repo)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent.parent


def sciezka():
    """Ścieżka pliku ze słownikiem zamian."""
    return _app_dir() / "config" / "zamiany_opisow.json"


# ------------------------------------------------------------------ klucz

def _podziel(wartosc):
    """Dzieli wartość na (część opisowa, część z powierzchnią).

    Rozdzielamy po OSTATNIM „|", bo opisy same mogą zawierać „|"
    (np. „Rola | droga polna|0.20" -> opis „Rola | droga polna").
    """
    v = (wartosc or "").strip()
    i = v.rfind("|")
    if i < 0:
        return v, ""
    return v[:i].strip(), v[i + 1:].strip()


def klucz(wartosc):
    """Zwraca klucz zamiany dla danej wartości (część opisowa)."""
    return _podziel(wartosc)[0]


# oznaczenie pola A2/A5 -> to samo, co klucz (alias dla czytelności w rdzeniu)
OPIS = klucz


# ------------------------------------------------------------------ plik

def wczytaj():
    """Wczytuje słownik zamian: {opis_wejsciowy: opis_docelowy}."""
    p = sciezka()
    try:
        if not p.exists():
            return {}
        raw = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    zam = raw.get("zamiany") if isinstance(raw, dict) else None
    if not isinstance(zam, dict):
        return {}
    # porządkujemy: same pary tekstowe, bez pustych kluczy
    czyste = {}
    for k, v in zam.items():
        k = (k or "").strip()
        v = (v or "").strip()
        if k and v and k != v:
            czyste[k] = v
    return czyste


def zapisz(slownik):
    """Zapisuje słownik zamian. Zwraca liczbę pozycji."""
    p = sciezka()
    p.parent.mkdir(parents=True, exist_ok=True)
    zam = {(k or "").strip(): (v or "").strip()
           for k, v in (slownik or {}).items()
           if (k or "").strip() and (v or "").strip()}
    dane = {"wersja": 1, "zamiany": zam}
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(dane, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)          # zapis atomowy — nie zostawia połowy pliku
    return len(zam)


# ------------------------------------------------------------------ logika

def zastosuj(wartosc, slownik=None):
    """Zamienia opis wg słownika, zachowując część z powierzchnią.

    Jeśli słownik nie zawiera klucza — zwraca wartość bez zmian.
    """
    if slownik is None:
        slownik = wczytaj()
    if not slownik:
        return wartosc
    opis, pow_ = _podziel(wartosc)
    if opis in slownik:
        nowy = slownik[opis]
        return ("%s|%s" % (nowy, pow_)) if pow_ else nowy
    return wartosc


def zapamietaj(pary, slownik=None):
    """Dodaje zamiany do słownika i zapisuje go.

    ``pary`` to lista par (oryginał, nowa_wartość) albo słownik
    {oryginał: nowa_wartość}. Klucze liczymy tak samo jak w ``zastosuj``.
    Zwraca (nowy_słownik, ile_dodano).
    """
    if slownik is None:
        slownik = wczytaj()
    if isinstance(pary, dict):
        pary = list(pary.items())
    dodano = 0
    for oryginal, nowy in (pary or []):
        k = klucz(oryginal)
        v, _ = _podziel(nowy)
        if not k or not v or k == v:
            continue
        if slownik.get(k) != v:
            slownik[k] = v
            dodano += 1
    zapisz(slownik)
    return slownik, dodano


def usun(klucze, slownik=None):
    """Usuwa podane klucze ze słownika i zapisuje. Zwraca ile usunięto."""
    if slownik is None:
        slownik = wczytaj()
    ile = 0
    for k in (klucze or []):
        if k in slownik:
            del slownik[k]
            ile += 1
    zapisz(slownik)
    return slownik, ile


def wyczysc():
    """Kasuje cały słownik zamian."""
    return zapisz({})
