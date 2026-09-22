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
        if not files:
            return 0

        count = 0
        total = len(files)
        self.start_progress_tracking(total, "Czyszczenie TXT")

        for idx, f in enumerate(files, start=1):
            self.check_stop()
            self.set_progress((idx - 1) / total if total else 1, current_file=f.name, current=idx - 1)
            rel_path = f.relative_to(in_dir)
            flat_rel_path = flatten_rel_path(rel_path)
            target = out_dir / flat_rel_path
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
            try:
                os.remove(bat_path)
            except:
                pass
            try:
                os.remove(log_path)
            except:
                pass

