#!/usr/bin/env python3
"""
Forestly — punkt wejścia interfejsu WEBOWEGO (PyWebView)
=================================================================

Uruchamia aplikację z nowoczesnym GUI opartym o HTML/CSS/JS (WebView2),
z całą logiką w Pythonie — tę samą, co wersja CustomTkinter (main.py).

  python main_web.py   →  aplikacja GUI (web)

Wymagania: pywebview (pip install pywebview).
Na Windows 10/11 WebView2 jest preinstalowany.
"""

import os
import sys
import warnings

warnings.filterwarnings("ignore", message=".*OLE2 inconsistency.*")
warnings.filterwarnings("ignore", message=".*file size.*not.*sector size.*")
warnings.filterwarnings("ignore", message=".*SSCS size.*")
import logging

logging.getLogger("xlrd").setLevel(logging.ERROR)
logging.getLogger("openpyxl").setLevel(logging.ERROR)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def get_index_html():
    """Ścieżka do webapp/index.html (działa też po zamrożeniu PyInstaller)."""
    base = os.path.dirname(os.path.abspath(__file__))
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "webapp", "index.html")
    return os.path.join(base, "webapp", "index.html")


def run_word_worker_cli():
    """Obsługa flagi --word-worker — uruchamia proces Word w tle.

    Identyczna logika jak w main.py; potrzebna, żeby EXE (PyInstaller,
    sys.executable = Forestly.exe) mógł sam pełnić rolę procesu Word COM.
    """
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
    # Tryb procesu w tle (m.in. EXE): Forestly.exe --word-worker IN OUT [opcje]
    if "--word-worker" in sys.argv:
        run_word_worker_cli()
        return

    import webview

    from app.config import kill_orphan_office_processes
    from app.gui.web_backend import WebBackend

    kill_orphan_office_processes()

    backend = WebBackend()

    window = webview.create_window(
        "Forestly",
        get_index_html(),
        js_api=backend,
        width=1520,
        height=960,
        min_size=(1100, 720),
        text_select=True,
    )
    webview.start()


if __name__ == "__main__":
    main()
