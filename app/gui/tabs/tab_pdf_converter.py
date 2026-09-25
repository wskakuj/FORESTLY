"""
Forestly — Mixin: TabPdfConverterMixin
"""

import customtkinter as ctk
from pathlib import Path
import threading
import traceback
import win32com.client
import pythoncom

from app.config import (
    is_file_locked,
)

class TabPdfConverterMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_pdf_converter_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        font_hint = ctk.CTkFont(family="Segoe UI", size=12)
        card_frame = ctk.CTkFrame(
            scroll_frame,
            fg_color="#252526",
            corner_radius=8,
            border_width=1,
            border_color="#333333",
        )
        card_frame.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="new")
        card_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            card_frame, text="Folder źródłowy:", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        self.pdfconv_source_entry = ctk.CTkEntry(
            card_frame,
            placeholder_text="Folder z plikami Office / PDF / obrazami...",
            height=36,
            border_width=1,
        )
        self.pdfconv_source_entry.grid(
            row=0, column=1, padx=5, pady=(20, 10), sticky="ew"
        )
        ctk.CTkButton(
            card_frame,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.pdfconv_source_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(20, 10))
        ctk.CTkLabel(
            card_frame,
            text="Folder docelowy PDF:",
            font=font_label,
            text_color="#E0E0E0",
        ).grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")
        self.pdfconv_output_entry = ctk.CTkEntry(
            card_frame,
            placeholder_text="Miejsce zapisu przekonwertowanych dokumentów...",
            height=36,
            border_width=1,
        )
        self.pdfconv_output_entry.grid(
            row=1, column=1, padx=5, pady=(0, 10), sticky="ew"
        )
        ctk.CTkButton(
            card_frame,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.pdfconv_output_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=(0, 10))
        info_text = "Obsługiwane formaty: DOC, DOCX, RTF, TXT, XLS, XLSX, CSV, JPG, PNG, BMP, TIF, WEBP"
        ctk.CTkLabel(
            card_frame, text=info_text, font=font_hint, text_color="#888888"
        ).grid(row=2, column=0, columnspan=3, padx=15, pady=(0, 20), sticky="w")
        self.pdfconv_start_btn = ctk.CTkButton(
            scroll_frame,
            text="Konwertuj wszystko do PDF",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=44,
            corner_radius=6,
            command=self.start_pdf_converter_pipeline,
        )
        self.pdfconv_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_pdf_converter_pipeline(self):
        source_folder = (
            self.pdfconv_source_entry.get().strip() if self.pdfconv_source_entry else ""
        )
        output_folder = (
            self.pdfconv_output_entry.get().strip() if self.pdfconv_output_entry else ""
        )
        pliki_drop = [x for x in (getattr(self, "_pc_pliki", None) or [])
                      if Path(x).is_file()]
        if not pliki_drop and (not source_folder or not Path(source_folder).exists()):
            self.log("[KONWERTER PDF] Wskaż folder źródłowy albo przeciągnij pliki.")
            self.update_status("Brak plików do konwersji", "#D83B01", animate=False)
            return
        if not output_folder:
            self.log("[KONWERTER PDF] Wskaż folder docelowy PDF.")
            self.update_status("Brak folderu docelowego", "#D83B01", animate=False)
            return
        if self.running:
            return
        self.last_output_dir = Path(output_folder)
        self._disable_ui_for_process()
        if pliki_drop:
            self.log(f"[KONWERTER PDF] Przeciągnięte pliki: {len(pliki_drop)}")
        if source_folder and Path(source_folder).exists():
            self.log(f"[KONWERTER PDF] Źródło: {source_folder}")
        self.set_progress(0)
        threading.Thread(
            target=self.run_pdf_converter_thread,
            args=(source_folder if source_folder and Path(source_folder).exists() else None,
                  output_folder, pliki_drop),
            daemon=True,
        ).start()

    def run_pdf_converter_thread(self, source_folder_str, output_folder_str,
                                 pliki_drop=None):
        pythoncom.CoInitialize()
        word, excel = None, None
        try:
            source_folder = Path(source_folder_str)
            output_folder = Path(output_folder_str)
            supported_exts = {
                ".doc",
                ".docx",
                ".rtf",
                ".txt",
                ".xls",
                ".xlsx",
                ".csv",
                ".jpg",
                ".jpeg",
                ".png",
                ".bmp",
                ".tif",
                ".tiff",
                ".gif",
                ".webp",
            }
            files = sorted(
                [
                    p
                    for p in (source_folder.rglob("*")
                              if source_folder else [])
                    if p.is_file() and p.suffix.lower() in supported_exts
                ]
            ) if source_folder else []
            for dp in (pliki_drop or []):
                if dp not in files:
                    files.append(dp)
            files.sort()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible, word.DisplayAlerts = False, 0
            word.Application.ScreenUpdating = False
            word.Options.BackgroundSave = False
            word.Options.CheckSpellingAsYouType = False
            word.Options.CheckGrammarAsYouType = False
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible, excel.DisplayAlerts = False, False
            total = len(files)
            self.start_progress_tracking(total, "Konwersja do PDF")

            for idx, file_path in enumerate(files, start=1):
                self.check_stop()
                self.progress_current_file = file_path.name
                if is_file_locked(file_path):
                    self.log(f"POMINIĘTO ZABLOKOWANY PLIK: {file_path.name}")
                    continue
                try:
                    if source_folder and file_path.is_relative_to(source_folder):
                        rel_path = file_path.relative_to(source_folder)
                        target_dir = output_folder / rel_path.parent
                    else:
                        target_dir = output_folder
                    target_dir.mkdir(parents=True, exist_ok=True)
                    pdf_path = target_dir / f"{file_path.stem}.pdf"
                    self.convert_office_to_pdf(file_path, pdf_path, word, excel)
                except Exception as e:
                    self.log(f"Błąd konwersji {file_path.name}: {e}")
                self.set_progress(idx / total)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            # pliki przeciągnięte obsługane — czyścimy kolejkę (i listę w GUI)
            self._pc_pliki = []
            self._emit({"type": "pdfconv_files", "files": []})
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            if word is not None:
                try:
                    word.Quit()
                except:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

