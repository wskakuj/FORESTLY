#!/usr/bin/env python3
"""
Forestly — Punkt wejścia
===================================
Uruchamia aplikację GUI. Cała logika znajduje się w pakiecie `app/`.

Podwójna osobowość — tak jak oryginalny guipia.py:
  1. python main.py                         → aplikacja GUI
  2. python main.py --word-worker IN OUT ... → proces w tle (Word COM worker)

Po skompilowaniu przez PyInstaller, ten plik staje się .exe.
"""

import sys
import os
import warnings

# Wyciszanie ostrzeżeń z bibliotek xlrd/openpyxl przy czytaniu starych plików .xls
# (niegroźne "OLE2 inconsistency" i "file size not multiple of sector size")
warnings.filterwarnings("ignore", message=".*OLE2 inconsistency.*")
warnings.filterwarnings("ignore", message=".*file size.*not.*sector size.*")
warnings.filterwarnings("ignore", message=".*SSCS size.*")
import logging
logging.getLogger("xlrd").setLevel(logging.ERROR)
logging.getLogger("openpyxl").setLevel(logging.ERROR)

# Dodaj katalog projektu do ścieżki Pythona, aby `app` był importowalny
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def run_word_worker_cli():
    """Obsługa flagi --word-worker — uruchamia proces Word w tle."""
    import json
    import traceback

    log_file_path = None
    if "--log-file" in sys.argv:
        l_idx = sys.argv.index("--log-file")
        log_file_path = sys.argv[l_idx + 1]

    if log_file_path:
        class FileLogger:
            def __init__(self, filename):
                self.filename = filename
                with open(self.filename, "w", encoding="utf-8") as f:
                    f.write("")

            def write(self, text):
                with open(self.filename, "a", encoding="utf-8") as f:
                    f.write(str(text))

            def flush(self):
                pass

        sys.stdout = sys.stderr = FileLogger(log_file_path)

    try:
        from app.core.word_worker import run_word_worker

        idx = sys.argv.index("--word-worker")
        in_dir = sys.argv[idx + 1]
        out_dir = sys.argv[idx + 2]
        remove_names = "--remove-names" in sys.argv

        file_filter = []
        idx_scan = 0
        margins_config = {}
        while idx_scan < len(sys.argv):
            if sys.argv[idx_scan] == "--filter" and idx_scan + 1 < len(sys.argv):
                file_filter.append(sys.argv[idx_scan + 1])
                idx_scan += 2
                continue
            if sys.argv[idx_scan] == "--margins-file" and idx_scan + 1 < len(sys.argv):
                try:
                    with open(sys.argv[idx_scan + 1], 'r', encoding='utf-8') as mf:
                        margins_config = json.load(mf)
                except Exception:
                    pass
                idx_scan += 2
                continue
            idx_scan += 1

        if not file_filter:
            file_filter = ["Wszystkie"]

        run_word_worker(in_dir, out_dir, remove_names, file_filter, margins_config)

    except Exception as e:
        import traceback
        if log_file_path:
            try:
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"\n[BŁĄD KRYTYCZNY PROCESU WORD]: {e}\n")
                    f.write(traceback.format_exc())
            except Exception:
                pass
        sys.exit(1)

    sys.exit(0)


def main():
    # Tryb procesu w tle: python main.py --word-worker IN OUT [opcje]
    if "--word-worker" in sys.argv:
        run_word_worker_cli()
        return

    # Tryb normalny: aplikacja GUI (WEB — wersja CustomTkinter wycofana)
    from main_web import main as web_main

    web_main()


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()
