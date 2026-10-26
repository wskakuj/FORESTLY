# -*- coding: utf-8 -*-
"""FORESTLY: spójne i czytelne zgłaszanie błędów z kontekstem przetwarzania.

Moduł jest niezależny od GUI, Office i bibliotek zewnętrznych.
Nie wysyła raportów przez sieć. Szczegółowy traceback zapisujemy lokalnie,
a komunikat w interfejsie zawiera etap, plik, przyczynę i wskazówkę.
"""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import traceback
from uuid import uuid4


def _wskazowka(exc):
    """Podpowiedź oparta na typie błędu, bez zgadywania faktycznej przyczyny."""
    if isinstance(exc, FileNotFoundError):
        return ("Sprawdź, czy wskazany plik lub folder istnieje i czy nie został "
                "przeniesiony. Następnie ponów operację.")
    if isinstance(exc, PermissionError):
        return ("Sprawdź uprawnienia do pliku i folderu. Jeśli plik jest otwarty "
                "w Wordzie, Excelu lub innym programie, zamknij go i spróbuj ponownie.")
    if isinstance(exc, (UnicodeError, UnicodeDecodeError, UnicodeEncodeError)):
        return ("Plik może mieć inne kodowanie znaków niż oczekiwane. "
                "Sprawdź format i kodowanie pliku wejściowego.")
    if isinstance(exc, ModuleNotFoundError):
        return ("Brakuje wymaganego składnika programu. Sprawdź instalację "
                "FORESTLY i kompletność pakietu.")
    if isinstance(exc, TimeoutError):
        return ("Operacja przekroczyła limit czasu. Sprawdź, czy Word, Excel "
                "lub przeglądarka nie przestały odpowiadać.")
    if isinstance(exc, (IsADirectoryError, NotADirectoryError)):
        return "Sprawdź, czy wybrano właściwy plik lub folder."
    if isinstance(exc, OSError):
        return ("Sprawdź dostępność dysku, wolne miejsce, uprawnienia "
                "i czy plik nie jest zablokowany przez inny program.")
    if isinstance(exc, ValueError):
        return ("Sprawdź dane wejściowe i format pliku. Szczegóły techniczne "
                "pomogą wskazać nieprawidłowe pole lub wartość.")
    return ("Sprawdź wskazany plik i etap. Jeśli problem się powtórzy, "
            "skorzystaj ze szczegółów technicznych zapisanych w raporcie.")


def zbuduj_raport(exc, *, etap, plik=None, zadanie=None, krytyczny=True):
    """Konstruuje strukturę gotową do JSON i wyświetlenia w PyWebView.

    Obiekt wyjątku bywa przechwycony poza blokiem except (np. w teście),
    więc traceback może być pusty. Nie dopisujemy zmyślonej lokalizacji.
    """
    if not isinstance(exc, BaseException):
        raise TypeError("Oczekiwano wyjątku")
    sciezka = str(plik) if plik is not None else ""
    if not sciezka:
        sciezka = str(getattr(exc, "filename", "") or "")
    etap = str(etap or "Nieustalony etap")
    typ = type(exc).__name__
    przyczyna = str(exc).strip() or typ
    slady = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    return {
        "id": uuid4().hex[:10],
        "czas": datetime.now().astimezone().isoformat(timespec="seconds"),
        "zadanie": str(zadanie or ""),
        "etap": etap,
        "plik": sciezka,
        "typ": typ,
        "przyczyna": przyczyna,
        "wskazowka": _wskazowka(exc),
        "krytyczny": bool(krytyczny),
        "szczegoly": slady,
    }


def opis_do_logu(raport):
    """Zwięzły wiersz do dziennika pracy FORESTLY."""
    status = "BŁĄD" if raport["krytyczny"] else "BŁĄD PLIKU"
    czesci = [f"[{status} {raport['id']}]",
              f"Etap: {raport['etap']}"]
    if raport.get("plik"):
        czesci.append(f"Plik: {raport['plik']}")
    czesci.append(f"Przyczyna: {raport['typ']}: {raport['przyczyna']}")
    czesci.append(f"Co zrobić: {raport['wskazowka']}")
    return " | ".join(czesci)


def zapisz_raport_jsonl(raport, folder_wynikow):
    """Zapisuje raport do podfolderu diagnostyki; nigdy niczego nie nadpisuje.

    Brak folderu wyników oznacza brak zapisu (np. gdy zadanie jeszcze
    nie zostało uruchomione). Zwraca ścieżkę lub None.
    """
    if not folder_wynikow:
        return None
    folder = Path(folder_wynikow) / "_FORESTLY_DIAGNOSTYKA"
    folder.mkdir(parents=True, exist_ok=True)
    sciezka = folder / "bledy.jsonl"
    tekst = json.dumps(raport, ensure_ascii=False)
    with sciezka.open("a", encoding="utf-8", newline="\n") as f:
        f.write(tekst + "\n")
    return sciezka
