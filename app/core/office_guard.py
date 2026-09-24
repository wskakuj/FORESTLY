# -*- coding: utf-8 -*-
"""Rejestr procesów Office (Word/Excel) uruchomionych przez Forestly.

Problem: przerwane zadanie (np. Pełny Automat) zostawiało w tle proces
WINWORD.EXE uruchomiony przez COM — trzymał blokady na plikach, przez
co nie dało się później usunąć folderów wynikowych. Word startowany
przez COM żyje POZA drzewem procesów workera, więc taskkill /T używany
przy przerywaniu zadania go nie dosięga.

Rozwiązanie:
  * każdy nasz Word/Excel (COM) zapisuje swój PID do rejestru w %TEMP%
    (plik przeżywa awaryjne ubicie workera z parenta),
  * po poprawnym zamknięciu PID znika z rejestru,
  * po przerwaniu zadania ubijane są dokładnie procesy z rejestru —
    z kontrolą nazwy obrazu (tasklist), więc Word/Excel otwarte przez
    użytkownika nigdy nie są dotykane (ich PID-ów w rejestrze nie ma).
"""

import ctypes
import json
import os
import subprocess
import tempfile
import threading
from pathlib import Path

_LOCK = threading.Lock()
_OFFICE_IMAGES = ("WINWORD.EXE", "EXCEL.EXE")


def _registry_path():
    return Path(tempfile.gettempdir()) / "forestly_office_pids.json"


def _load():
    try:
        data = json.loads(_registry_path().read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save(pids):
    tmp = _registry_path().with_suffix(".tmp")
    tmp.write_text(json.dumps(sorted(pids)), encoding="utf-8")
    os.replace(str(tmp), str(_registry_path()))


def pid_of_app(app):
    """PID procesu aplikacji COM (Word/Excel mają właściwość Hwnd).

    Zwraca None, gdy PID-u nie da się ustalić (np. brak Hwnd).
    """
    try:
        hwnd = int(app.Hwnd)
    except Exception:
        return None
    if not hwnd:
        return None
    try:
        pid = ctypes.c_ulong()
        if ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)):
            return int(pid.value) or None
    except Exception:
        pass
    return None


def register(app):
    """Rejestruje proces aplikacji Office uruchomionej przez Forestly."""
    pid = pid_of_app(app)
    if pid is None:
        return None
    with _LOCK:
        pids = _load()
        if pid not in pids:
            pids.append(pid)
            _save(pids)
    return pid


def unregister(pid):
    """Usuwa PID z rejestru (po poprawnym zamknięciu aplikacji)."""
    if pid is None:
        return
    with _LOCK:
        pids = _load()
        if pid in pids:
            pids.remove(pid)
            _save(pids)


def _image_name(pid):
    """Nazwa obrazu procesu (np. WINWORD.EXE); None gdy nie żyje / nie Office."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "PID eq %d" % pid, "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10).stdout or ""
    except Exception:
        return None
    out = out.strip().upper()
    for img in _OFFICE_IMAGES:
        if img in out:
            return img
    return None


def kill_registered(log=None):
    """Ubija procesy Office zostawione w tle przez przerwane zadanie.

    Dotyka wyłącznie procesów uprzednio zarejestrowanych przez Forestly
    (i tylko gdy ich obraz to WINWORD.EXE/EXCEL.EXE — ochrona przed
    ponownym użyciem PID-a). Zwraca listę ubitych PID-ów.
    """
    with _LOCK:
        pids = _load()
        if not pids:
            return []
        _save([])  # rejestr czyszczony od razu — bez podwójnego ubijania
    killed = []
    for pid in pids:
        if _image_name(pid) is None:
            continue  # proces już nie żyje albo to nie Office — nie ruszaj
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=15)
            killed.append(pid)
        except Exception:
            pass
    if killed and log:
        try:
            log("[PORZĄDKI] Zamknięto proces(y) Office zostawione w tle: "
                + ", ".join(str(p) for p in killed))
        except Exception:
            pass
    return killed
