#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Forestly — pomocnik wydawania wersji (release)
=============================================
Jednym poleceniem zapisuje zmienione pliki do repo GitHub i wypuszcza
nowy release — Actions zbudują Forestly.exe i Forestly_OLD.exe,
a changelog trafi do opisu Release (okno „Co nowego" w programie).

Użycie (w folderze repo FORESTLY):
    python release.py        → kreator krok po kroku
    python release.py -k     → bez pytania o potwierdzenie (konto gotowe)

Wymagania: git (zalogowany — klon robiony przez HTTPS z zapamiętanym hasłem).
"""

import re
import subprocess
import sys
import webbrowser
from pathlib import Path

REPO = Path(__file__).resolve().parent
CONFIG = REPO / "app" / "config.py"
NOTES = REPO / "RELEASE_NOTES.md"
GITHUB_URL = "https://github.com/wskakuj/FORESTLY"


def git(*args, check=True):
    """Uruchamia git w katalogu repo, zwraca stdout (lub exits przy błędzie)."""
    r = subprocess.run(["git", "-C", str(REPO), *args],
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if check and r.returncode != 0:
        print("\n✗ BŁĄD git " + " ".join(args))
        if r.stdout.strip():
            print(r.stdout.strip())
        if r.stderr.strip():
            print(r.stderr.strip())
        print("\nNic nie wysłano — popraw problem i uruchom release.py ponownie.")
        sys.exit(1)
    return r.stdout.strip()


def read_current_version():
    m = re.search(r'CURRENT_VERSION = "([^"]+)"',
                  CONFIG.read_text(encoding="utf-8"))
    if not m:
        print("✗ Nie znaleziono CURRENT_VERSION w app/config.py")
        sys.exit(1)
    return m.group(1)


def set_current_version(ver):
    s = CONFIG.read_text(encoding="utf-8")
    s2 = re.sub(r'CURRENT_VERSION = "[^"]+"',
                f'CURRENT_VERSION = "{ver}"', s, count=1)
    CONFIG.write_text(s2, encoding="utf-8")


def next_patch(v):
    m = re.match(r"^v(\d+)\.(\d+)\.(\d+)$", v)
    if not m:
        return None
    a, b, c = (int(x) for x in m.groups())
    return f"v{a}.{b}.{c + 1}"


def sanitize_notes(text):
    """Usuwa z changelogu encje HTML i gwiazdki markdownu, które w oknie
    „Co nowego" w programie wyglądałyby jak krzaczki."""
    A = "&"
    pairs = (
        (A + "amp;#x20;", " "),
        (A + "amp;nbsp;", " "),
        (A + "#x20;", " "),
        (A + "nbsp;", " "),
        (A + "#160;", " "),
        (A + "#xa0;", " "),
        (A + "quot;", '"'),
        (A + "#39;", "'"),
        (A + "lt;", "<"),
        (A + "gt;", ">"),
        (A + "amp;", A),
    )
    for _pass in range(2):   # dwa przebiegi — na wypadek podwójnych encji
        for old, new in pairs:
            text = text.replace(old, new)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    # wypunktowanie gwiazdką → myślnik (poprawne kropki na GitHubie)
    text = re.sub(r"(?m)^\s*\*\s+", "- ", text)
    return text

def edit_changelog(ver):
    """Changelog: notepad na Windows, wpisywanie w konsoli gdzie indziej."""
    header = f"# Co nowego w {ver}\n\n"
    if sys.platform == "win32":
        NOTES.write_text(header + "- \n", encoding="utf-8")
        print("\nOtwieram Notatnik — napisz changelog, ZAPISZ i zamknij okno.")
        try:
            subprocess.run(["notepad.exe", str(NOTES)], check=False)
        except FileNotFoundError:
            pass
        raw = NOTES.read_text(encoding="utf-8")
        clean = sanitize_notes(raw)
        if clean != raw:
            NOTES.write_text(clean, encoding="utf-8")
            print("   (wyczyściłem znaki specjalne, które psułyby okno Co nowego)")
        body = clean.strip()
        if body in (header.strip(), header.strip() + "-"):
            print("   (changelog pusty — użyję tylko listy commitów z GitHuba)")
            return
        return
    # wariant konsolowy (test / inne systemy)
    print("\nWpisuj linie changelogu; pusta linia kończy:")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if not line.strip():
            break
        lines.append(line)
    NOTES.write_text(sanitize_notes(header + "\n".join(lines)) + "\n",
                     encoding="utf-8")


def main():
    print("=" * 62)
    print("  FORESTLY — wydawanie nowej wersji")
    print("=" * 62)

    # 0) czy to w ogóle repo gita?
    git("rev-parse", "--verify", "HEAD")

    # 1) co się zmieniło? (pliki + ewentualne commity czekające na wysłanie)
    status = git("status", "--short")
    unpushed = git("log", "--branches", "--not", "--remotes", "--oneline", check=False)
    if not status and not unpushed:
        print("\nBrak zmian — drzewo robocze czyste. Nie ma czego wydawać.")
        sys.exit(0)
    if status:
        print(f"\nZmienione / nowe pliki ({len(status.splitlines())}):")
        for line in status.splitlines():
            print("   " + line)
    if unpushed:
        print("\nUwaga: są już commity niewysłane na GitHub —")
        print("wydanie dokończy ich wysyłkę.")

    # 2) nowa wersja
    cur = read_current_version()
    prop = next_patch(cur) or "v0.0.1"
    print(f"\nAktualna wersja (app/config.py): {cur}")
    try:
        ans = input(f"Nowa wersja [{prop}]: ").strip() or prop
    except EOFError:
        ans = prop
    if not re.match(r"^v\d+\.\d+\.\d+$", ans):
        print("✗ Wersja musi być w formacie vX.Y.Z (np. v2.0.2)")
        sys.exit(1)
    if git("tag", "-l", ans):
        print(f"✗ Tag {ans} już istnieje w repo — wybierz inny numer.")
        sys.exit(1)

    # 3) opis commita
    try:
        msg = input(f"Krótki opis zmian [Wersja {ans}]: ").strip() or f"Wersja {ans}"
    except EOFError:
        msg = f"Wersja {ans}"

    # 4) changelog
    edit_changelog(ans)

    # 5) potwierdzenie
    print("\n" + "-" * 62)
    print(f"Wersja : {ans}   (obecnie: {cur})")
    print(f"Commit : {msg}")
    if NOTES.exists():
        preview = [l for l in NOTES.read_text(encoding="utf-8").splitlines() if l.strip()]
        print("Release:")
        for l in preview[:5]:
            print("   " + l)
        if len(preview) > 5:
            print(f"   … (+{len(preview) - 5} linii)")
    print("-" * 62)
    if "-k" not in sys.argv:
        try:
            ok = input("\nWypuścić wersję? [t/N]: ").strip().lower()
        except EOFError:
            ok = "n"
        if ok not in ("t", "tak", "y", "yes"):
            print("Anulowano — nic nie wysłano.")
            sys.exit(0)

    # 6) wykonanie
    print("\nUstawiam wersję w app/config.py…")
    set_current_version(ans)
    print("Zapisuję pliki (git add + commit)…")
    git("add", "-A")
    staged = git("diff", "--cached", "--name-only")
    if staged:
        git("commit", "-m", msg)
    print("Wysyłam zmiany na GitHub (push)…")
    git("push")
    print(f"Taguję {ans} i wysyłam tag — Actions budują EXE…")
    git("tag", ans)
    git("push", "origin", ans)

    print("\n✓ WYPUŚCZONO WERSJĘ " + ans)
    print(f"  Postęp budowy : {GITHUB_URL}/actions")
    print(f"  Release (po paru minutach): {GITHUB_URL}/releases")
    if "--open" in sys.argv:
        webbrowser.open(f"{GITHUB_URL}/actions")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrzerwano — nic nie wysłano.")
