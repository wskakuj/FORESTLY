"""
Forestly — Mixin: TabWordMixin
"""

import customtkinter as ctk
import time
import json
import os
import sys
import subprocess
import tempfile

from app.config import (
    SEQUENCES_TO_REMOVE, add_tooltip, flatten_rel_path, normalize_filter_selection,
)

class TabWordMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def _setup_word_extras(self, card_frame, row_idx):
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        ctk.CTkLabel(
            card_frame, text="Konwertuj tylko:", font=font_label, text_color="#E0E0E0"
        ).grid(row=row_idx, column=0, padx=15, pady=(0, 20), sticky="ne")
        options_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        options_frame.grid(
            row=row_idx, column=1, columnspan=2, padx=5, pady=(0, 20), sticky="w"
        )
        choices = [
            "Wszystkie",
            "REJESTR1",
            "OPTAX",
            "TAB_KLW3",
            "WSKAZ1",
            "HALIZNY",
            "WYK_NEG",
            "OPIS",
            "ZEST1",
            "WK_ZM1",
        ]
        self.word_filter_vars = {}
        self.word_filter_checkboxes = {}
        for idx, choice in enumerate(choices):
            var = ctk.BooleanVar(value=(choice == "Wszystkie"))
            cb = ctk.CTkCheckBox(
                options_frame,
                text=choice,
                variable=var,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                fg_color="#0067C0",
                hover_color="#005A9E",
                command=lambda c=choice: self.on_word_filter_change(c),
            )
            cb.grid(row=idx // 3, column=idx % 3, padx=(0, 14), pady=4, sticky="w")
            self.word_filter_vars[choice] = var
            self.word_filter_checkboxes[choice] = cb
        add_tooltip(
            options_frame,
            "Możesz zaznaczyć wiele typów plików naraz. Opcja 'Wszystkie' wyklucza pozostałe.",
        )

        # --- DODANA TABELA MARGINESÓW ---
        self._build_margins_ui(card_frame, row_idx + 1, "WORD")

    def on_word_filter_change(self, changed_option):
        if not getattr(self, "word_filter_vars", None):
            return
        if changed_option == "Wszystkie":
            if self.word_filter_vars["Wszystkie"].get():
                for name, var in self.word_filter_vars.items():
                    if name != "Wszystkie":
                        var.set(False)
            else:
                if not any(
                        var.get()
                        for name, var in self.word_filter_vars.items()
                        if name != "Wszystkie"
                ):
                    self.word_filter_vars["Wszystkie"].set(True)
        else:
            if self.word_filter_vars[changed_option].get():
                self.word_filter_vars["Wszystkie"].set(False)
            else:
                if not any(
                        var.get()
                        for name, var in self.word_filter_vars.items()
                        if name != "Wszystkie"
                ):
                    self.word_filter_vars["Wszystkie"].set(True)

    def get_selected_word_filters(self):
        if not getattr(self, "word_filter_vars", None):
            return ["Wszystkie"]
        selected = [name for name, var in self.word_filter_vars.items() if var.get()]
        if not selected:
            return ["Wszystkie"]
        if "Wszystkie" in selected:
            return ["Wszystkie"]
        return selected

    def task_clean_txt(self, in_dir, out_dir, file_filter=None):
        files = list(in_dir.rglob("*.txt"))
        selected_filters = normalize_filter_selection(file_filter)
        if "WSZYSTKIE" not in selected_filters:
            files = [f for f in files if f.stem.upper() in selected_filters]

        # Nie wciągamy plików z folderów WYNIKOWYCH poprzednich przebiegów,
        # które leżą w drzewie źródłowym (np. stare 'TXT', 'PDF', 'Word') —
        # inaczej trafiają do wydruków jako dodatkowy folder "wsi".
        _WYNIKOWE = {"TXT", "WORD", "PDF", "PDF POLACZONE",
                     "PDF BEZ PUSTYCH STRON", "Z NAZWISKAMI", "BEZ NAZWISK"}
        _pomijane = [f for f in files
                     if any(_p.upper() in _WYNIKOWE
                            for _p in f.relative_to(in_dir).parts[:-1])]
        if _pomijane:
            self.log(f"[TXT] Pomijam {len(_pomijane)} plik(ów) z folderów "
                     f"wynikowych (TXT/PDF/Word) znalezionych w źródle.")
            files = [f for f in files if f not in set(_pomijane)]

        # Struktura mietka to <wieś>/PLIK.TXT albo <wieś>/<obręb>.001/PLIK.TXT
        # (ew. PLIK.TXT, gdy wskazano samą wieś / obręb). TXT-y zagnieżdżone
        # głębiej to zapasowe kopie po starych biegach (np. "<wieś>/Nowy
        # folder/") — nie mogą wejść do wydruków jako osobny pakiet "wsi".
        _glebokie = []
        for _f in files:
            _czesci = _f.relative_to(in_dir).parts
            _ok = (len(_czesci) <= 2
                   or (len(_czesci) == 3
                       and (_czesci[1].upper().startswith("WOL")
                            or _czesci[1].lower().endswith(".001"))))
            if not _ok:
                _glebokie.append(_f)
        if _glebokie:
            _przykl = ", ".join(sorted({f.parent.name for f in _glebokie})[:4])
            self.log(f"[TXT] Pomijam {len(_glebokie)} plik(ów) TXT leżących "
                     f"głębiej niż folder wsi/.001 (folder {_przykl}) — "
                     f"zapasowe kopie po starych biegach.")
            files = [f for f in files if f not in set(_glebokie)]
        if not files:
            return 0

        # Zapamiętujemy ORYGINALNE położenie każdego TXT-a — po biegu
        # wyczyszczone pliki wracają DOKŁADNIE tam, skąd przyszły
        # (czyli do folderu .001 obok DBF-ów), a nie obok niego.
        txt_map = {}
        self._txt_map = txt_map

        count = 0
        total = len(files)
        self.start_progress_tracking(total, "Czyszczenie TXT")

        for idx, f in enumerate(files, start=1):
            self.check_stop()
            self.set_progress((idx - 1) / total if total else 1, current_file=f.name, current=idx - 1)
            rel_path = f.relative_to(in_dir)
            flat_rel_path = flatten_rel_path(rel_path)
            target = out_dir / flat_rel_path
            from pathlib import Path as _P
            klucz = os.path.abspath(str(target))
            stary = txt_map.get(klucz)
            # gdy ten sam wydruk leży i w .001, i w folderze wsi — wraca do .001
            def _w001(sciezka):
                return any(c.lower().endswith(".001")
                           for c in _P(sciezka).parts)
            if stary is None or (_w001(f) and not _w001(stary)):
                txt_map[klucz] = os.path.abspath(str(f))
            target.parent.mkdir(parents=True, exist_ok=True)
            try:
                with open(f, "rb") as file:
                    content = file.read()
                for seq in SEQUENCES_TO_REMOVE:
                    content = content.replace(seq, b"")
                with open(target, "wb") as file:
                    file.write(content)
                count += 1
                self.set_progress(idx / total if total else 1, current_file=f.name, current=idx)
            except Exception as e:
                self.log(f"Błąd pliku {f.name}: {e}")

        # TUTAJ BYŁ BŁĄD - to musi być na równi z "for", a nie wewnątrz niego!
        return count

    def task_word_processing_subprocess(
            self, in_dir, out_dir, remove_names, file_filter=None, margins_dict=None
    ):
        # main.py obsługuje flagę --word-worker (podwójna osobowość, jak oryginalny guipia.py)
        # Znajdź main.py względem tego pliku: app/gui/tabs/ → ../../../main.py
        python_exe = sys.executable
        frozen = bool(getattr(sys, "frozen", False))
        # EXE (PyInstaller): worker uruchamia ten sam plik .exe
        worker_script = None if frozen else os.path.abspath(
            os.path.join(os.path.dirname(__file__), "..", "..", "..", "main.py")
        )
        remove_flag = " --remove-names" if remove_names else ""
        selected_filters = normalize_filter_selection(file_filter)
        filter_flag = (
            ""
            if "WSZYSTKIE" in selected_filters
            else "".join(f' --filter "{flt}"' for flt in sorted(selected_filters))
        )

        # Zapis margins_dict do tymczasowego JSONa
        margins_file_path = ""
        if margins_dict:
            m_fd, margins_file_path = tempfile.mkstemp(suffix=".json")
            os.close(m_fd)
            with open(margins_file_path, "w", encoding="utf-8") as mf:
                json.dump(margins_dict, mf)
        margins_flag = f' --margins-file "{margins_file_path}"' if margins_file_path else ""

        log_fd, log_path = tempfile.mkstemp(suffix=".log")
        os.close(log_fd)

        if frozen:
            cmd_base = f'"{python_exe}" --word-worker "{str(in_dir).rstrip(r"/")}" "{str(out_dir).rstrip(r"/")}" --log-file "{log_path}"'
        else:
            cmd_base = f'"{python_exe}" -u "{worker_script}" --word-worker "{str(in_dir).rstrip(r"/")}" "{str(out_dir).rstrip(r"/")}" --log-file "{log_path}"' 
        bat_content = f"@echo off\nchcp 65001 >nul\nset PYTHONIOENCODING=utf-8\nset PYTHONUNBUFFERED=1\n{cmd_base}{remove_flag}{filter_flag}{margins_flag}\nexit /b %errorlevel%\n"
        with tempfile.NamedTemporaryFile(
                "w", suffix=".bat", delete=False, encoding="utf-8"
        ) as bat_file:
            bat_file.write(bat_content)
        bat_path = bat_file.name
        try:
            process = subprocess.Popen(
                ["cmd", "/c", bat_path], creationflags=subprocess.CREATE_NO_WINDOW
            )
            with open(log_path, "r", encoding="utf-8") as f:
                while True:
                    if self.stop_event.is_set():
                        try:
                            subprocess.run(
                                ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL,
                            )
                        except Exception:
                            process.kill()
                        self.log(
                            "Proces ukrytego Worda (MIETEK) zablokowany i ugaszony z powodzeniem."
                        )
                        raise InterruptedError()
                    line = f.readline()
                    if line:
                        self.log(line.rstrip())
                    elif process.poll() is not None:
                        for remaining_line in f.readlines():
                            if remaining_line:
                                self.log(remaining_line.rstrip())
                        break
                    else:
                        time.sleep(0.1)
        finally:
            if process.poll() is None:
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                except:
                    process.kill()
            # Word startowany przez COM żyje POZA drzewem procesów workera
            # i przeżywa taskkill /T — ubijamy go z rejestru, bo inaczej
            # trzyma blokady na folderach wynikowych po przerwanym zadaniu
            try:
                from app.core import office_guard
                office_guard.kill_registered(log=self.log)
            except Exception:
                pass
            try:
                os.remove(bat_path)
            except:
                pass
            try:
                os.remove(log_path)
            except:
                pass

