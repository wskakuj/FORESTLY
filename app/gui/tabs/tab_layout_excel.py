"""
Forestly — Mixin: TabLayoutExcelMixin
"""

import customtkinter as ctk
from pathlib import Path
import threading
import re
import traceback
from PIL import Image
from pypdf import PdfWriter
from pypdf import PdfReader
import win32com.client
import pythoncom

from app.config import (
    is_file_locked,
)

class TabLayoutExcelMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_layout_excel_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        card = ctk.CTkFrame(
            scroll_frame,
            fg_color="#252526",
            corner_radius=8,
            border_width=1,
            border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            card, text="Folder STR_TYT:", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.layout_title_folder_entry = ctk.CTkEntry(
            card,
            placeholder_text="Folder ze stronami tytułowymi STRTYT*.docx",
            height=36,
        )
        self.layout_title_folder_entry.grid(
            row=0, column=1, padx=5, pady=(15, 8), sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.layout_title_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card, text="Folder z opisami:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=8, sticky="w")
        self.layout_opisy_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder z plikami opisów", height=36
        )
        self.layout_opisy_folder_entry.grid(
            row=1, column=1, padx=5, pady=8, sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.layout_opisy_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=8)
        ctk.CTkLabel(
            card, text="Folder z raportami:", font=font_label, text_color="#E0E0E0"
        ).grid(row=2, column=0, padx=15, pady=8, sticky="w")
        self.layout_raporty_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder z raportami Excel", height=36
        )
        self.layout_raporty_folder_entry.grid(
            row=2, column=1, padx=5, pady=8, sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.layout_raporty_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=2, column=2, padx=15, pady=8)
        ctk.CTkLabel(
            card, text="Folder docelowy PDF:", font=font_label, text_color="#E0E0E0"
        ).grid(row=3, column=0, padx=15, pady=(8, 15), sticky="w")
        self.layout_output_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder wyjściowy na gotowe PDF", height=36
        )
        self.layout_output_folder_entry.grid(
            row=3, column=1, padx=5, pady=(8, 15), sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.layout_output_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=3, column=2, padx=15, pady=(8, 15))
        self.layout_merge_btn = ctk.CTkButton(
            scroll_frame,
            text="Twórz gotowe PDF",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=44,
            corner_radius=6,
            command=self.start_layout_excel_pipeline,
        )
        self.layout_merge_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_layout_excel_pipeline(self):
        title_folder = (
            self.layout_title_folder_entry.get().strip()
            if self.layout_title_folder_entry
            else ""
        )
        opisy_folder = (
            self.layout_opisy_folder_entry.get().strip()
            if self.layout_opisy_folder_entry
            else ""
        )
        raporty_folder = (
            self.layout_raporty_folder_entry.get().strip()
            if self.layout_raporty_folder_entry
            else ""
        )
        output_folder = (
            self.layout_output_folder_entry.get().strip()
            if self.layout_output_folder_entry
            else ""
        )
        if (
                not title_folder
                or not opisy_folder
                or not raporty_folder
                or not output_folder
        ):
            return
        if self.running:
            return
        self.last_output_dir = Path(output_folder)
        self._disable_ui_for_process()
        self.log(f"[WYŁOŻENIE EXCEL] Procedura tworzenia gotowych paczek w toku...")
        self.set_progress(0)
        threading.Thread(
            target=self.run_layout_excel_thread,
            args=(title_folder, opisy_folder, raporty_folder, output_folder),
            daemon=True,
        ).start()

    def run_layout_excel_thread(
            self, title_folder_str, opisy_folder_str, raporty_folder_str, output_folder_str
    ):
        pythoncom.CoInitialize()
        word, excel = None, None
        _word_pid = _excel_pid = None
        try:
            # --- SŁOWNIK ZAKŁADEK DLA ARKUSZY EXCEL ---
            SHEET_BOOKMARKS = {
                "TPM_FL": "Zestawienie powierzchni i miąższości gatunków panujących w klasach i podklasach wieku według głównych funkcji lasu",
                "TPM_TH": "Zestawienie Powierzchni I Miąższości Gatunków Panujących W Typach Siedliskowych Lasu Wg. Klas I Podklas Wieku",
                "Zestawienie": "Zestawienie zadań gospodarczych projektowanych do wykonania",
                "WykazPow": "Wykaz powierzchni leśnych niezalesionych",
                "OT": "Opis Taksacyjny",
                "WykazWlasc": "Wykaz właścicieli",
                "REJ": "Rejestr działek leśnych i gruntów do zalesienia wg właścicieli",
                "Sheet4": "Rejestr działek leśnych i gruntów do zalesienia wg właścicieli",
                "WykazDzialek": "Wykaz działek",
                "Skroty": "Wykaz skrótów i symboli"
            }
            # ------------------------------------------

            # --- ZBIERAMY WSZYSTKIE WSIE (Z 3 FOLDERÓW) ---
            all_villages = self._get_all_villages(title_folder_str, opisy_folder_str, raporty_folder_str)
            if not all_villages:
                raise Exception("Nie odnaleziono żadnych plików wsi w podanych folderach.")

            # --- KONTROLA KOMPLETNOŚCI WYŁOŻENIA ---
            warnings = []
            for village_name in all_villages:
                path_title = self.find_matching_file(Path(title_folder_str), village_name)
                path_opis = self.find_matching_file(Path(opisy_folder_str), village_name)
                path_raport = self.find_matching_file(Path(raporty_folder_str), village_name)

                missing = []
                if not path_title: missing.append("STR_TYT")
                if not path_opis: missing.append("OPIS")
                if not path_raport: missing.append("RAPORT / REJESTR")

                if missing:
                    warnings.append(f"• Wieś {village_name}: brak -> {', '.join(missing)}")

            if warnings:
                self.log("[KONTROLA] Wykryto braki w plikach wyłożenia. Oczekiwanie na decyzję...")
                if not self.show_validation_window_sync("Wykryto brakujące części w procedurze WYŁOŻENIA:", warnings):
                    raise InterruptedError("Operacja Wyłożenia Excel przerwana przez użytkownika.")
            # ----------------------------------------

            temp_folder = Path(output_folder_str) / "_TEMP_PDF_WYLOZENIE"
            temp_folder.mkdir(parents=True, exist_ok=True)
            word = win32com.client.DispatchEx("Word.Application")
            try:
                from app.core import office_guard
                _word_pid = office_guard.register(word)
            except Exception:
                pass
            word.Visible, word.DisplayAlerts = False, 0
            word.Application.ScreenUpdating = False
            word.Options.BackgroundSave = False
            word.Options.CheckSpellingAsYouType = False
            word.Options.CheckGrammarAsYouType = False
            excel = win32com.client.DispatchEx("Excel.Application")
            try:
                _excel_pid = office_guard.register(excel)
            except Exception:
                pass
            excel.Visible, excel.DisplayAlerts = False, False

            total = len(all_villages)
            self.start_progress_tracking(total, "Wyłożenie Excel")

            for idx, village_name in enumerate(all_villages, start=1):
                self.check_stop()
                self.progress_current_file = village_name
                try:
                    # pdf_files to lista tupli: (ścieżka_pdf, nazwa_zakładki_w_drzewku)
                    pdf_files = []

                    path_title = self.find_matching_file(Path(title_folder_str), village_name)
                    if path_title and not is_file_locked(path_title):
                        pdf_str = temp_folder / f"1_STR_{village_name}.pdf"
                        if self.convert_office_to_pdf(path_title, pdf_str, word, excel):
                            pdf_files.append((pdf_str, "Strona tytułowa"))

                    path_opis = self.find_matching_file(Path(opisy_folder_str), village_name)
                    if path_opis and not is_file_locked(path_opis):
                        pdf_opis = temp_folder / f"2_OPIS_{village_name}.pdf"
                        if self.convert_office_to_pdf(path_opis, pdf_opis, word, excel):
                            pdf_files.append((pdf_opis, "Opis ogólny"))

                    path_raport = self.find_matching_file(Path(raporty_folder_str), village_name)
                    if path_raport and not is_file_locked(path_raport):
                        wb = None
                        try:
                            # Otwieramy skoroszyt Excela
                            wb = excel.Workbooks.Open(str(path_raport))
                            # Przechodzimy przez KAŻDY arkusz osobno
                            for ws_idx in range(1, wb.Worksheets.Count + 1):
                                if self.stop_event.is_set():
                                    raise InterruptedError()
                                ws = wb.Worksheets(ws_idx)

                                safe_ws_name = "".join(
                                    c for c in ws.Name if c.isalnum() or c in (" ", "_", "-")).strip()
                                pdf_ws = temp_folder / f"3_RAPORT_{village_name}_{safe_ws_name}.pdf"

                                try:
                                    # Eksportujemy tylko dany arkusz do pojedynczego PDF
                                    ws.ExportAsFixedFormat(0, str(pdf_ws))

                                    # Uodparniamy na ukryte spacje i wielkość liter w nazwie arkusza Excel
                                    sheet_name_clean = ws.Name.strip().upper()

                                    # Tworzymy w locie słownik z kluczami pisanymi wyłącznie dużymi literami
                                    bookmarks_upper = {k.strip().upper(): v for k, v in SHEET_BOOKMARKS.items()}

                                    # Pobieramy pełną nazwę zakładki
                                    bookmark_label = bookmarks_upper.get(sheet_name_clean, ws.Name)
                                    pdf_files.append((pdf_ws, bookmark_label))
                                except Exception as e:
                                    self.log(f"  [Ostrzeżenie] Pominięto arkusz {ws.Name}: {e}")

                        except InterruptedError:
                            raise
                        except Exception as e:
                            self.log(f"  [Błąd] Problem z arkuszami w pliku {path_raport.name}: {e}")
                        finally:
                            if wb is not None:
                                wb.Close(False)

                    if pdf_files:
                        writer = PdfWriter()
                        current_page = 0
                        for pdf_path, bookmark_label in pdf_files:
                            reader = PdfReader(str(pdf_path))
                            num_pages = len(reader.pages)

                            # Dodajemy czyste strony
                            for page in reader.pages:
                                writer.add_page(page)

                            writer.add_outline_item(bookmark_label, current_page)
                            current_page += num_pages

                        # --- NOWE: WSTRZYKIWANIE METADANYCH ---
                        writer.add_metadata({
                            "/Title": f"UPUL - {village_name.upper()}",
                            "/Author": "Agencja Cezar",
                            "/Creator": "Forestly",
                            "/Producer": "Forestly"
                        })
                        # --------------------------------------

                        with open(Path(output_folder_str) / f"Gotowy_{village_name}.pdf", "wb") as out_f:
                            writer.write(out_f)
                except Exception as e:
                    self.log(f"Błąd dla wsi {village_name}: {e}")
                self.set_progress(idx / total)

            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
        except InterruptedError as ie:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log(f"\n{ie}")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            if word:
                try:
                    word.Quit()
                except:
                    pass
                try:
                    office_guard.unregister(_word_pid)
                except Exception:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
                try:
                    office_guard.unregister(_excel_pid)
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

    def find_matching_file(self, folder_path, village_name):
        if not folder_path.exists():
            return None
        candidates = sorted(
            [
                p
                for p in folder_path.iterdir()
                if p.is_file() and not p.name.startswith("~$")
            ]
        )

        # Agresywna normalizacja - usuwamy wszystkie spacje, myślniki i podkreślniki
        v_norm = re.sub(r'[\s_\-]', '', village_name.lower())

        for file_path in candidates:
            # Normalizujemy tak samo nazwę pliku, który sprawdzamy
            c_norm = re.sub(r'[\s_\-]', '', file_path.stem.lower())

            # Sprawdzamy czy znormalizowana nazwa wsi zawiera się w znormalizowanej nazwie pliku
            if v_norm and v_norm in c_norm:
                return file_path

        return None

    def _get_all_villages(self, title_folder_str, opisy_folder_str, raporty_folder_str):
        """Zbiera unikalne nazwy wsi ze wszystkich trzech folderów źródłowych."""
        all_villages = set()

        # 1. Szukamy wsi w plikach STR_TYT
        if Path(title_folder_str).exists():
            for p in Path(title_folder_str).iterdir():
                if p.is_file() and p.name.lower().startswith("str_tyt_"):
                    all_villages.add(p.stem[8:].strip().upper())

        # 2. Szukamy wsi w plikach raportów Excel
        if Path(raporty_folder_str).exists():
            for p in Path(raporty_folder_str).iterdir():
                if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"} and not p.name.startswith("~$"):
                    v = self.extract_village_name_from_excel(p.name)
                    if v:
                        all_villages.add(v.strip().upper())

        # 3. Szukamy wsi w folderze Opisów
        if Path(opisy_folder_str).exists():
            for p in Path(opisy_folder_str).iterdir():
                if p.is_file() and not p.name.startswith("~$"):
                    name = p.stem.upper()
                    # Wycinamy wszystko typu "OPIS_", "OPIS OG_", "OPIS OGOLNY_" itd.
                    name = re.sub(r"^OPIS[\s_]*(OG[\w]*|OGÓLNY)?[\s_]*", "", name).strip()
                    if name:
                        all_villages.add(name)

        return sorted(list({v for v in all_villages if v}))

    def convert_office_to_pdf(self, input_path, output_path, word_app, excel_app):
        input_path = Path(input_path)
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        ext = input_path.suffix.lower()
        word_exts = {
            ".doc",
            ".docx",
            ".docm",
            ".dot",
            ".dotx",
            ".dotm",
            ".rtf",
            ".txt",
            ".odt",
        }
        excel_exts = {
            ".xls",
            ".xlsx",
            ".xlsm",
            ".xlsb",
            ".xlt",
            ".xltx",
            ".xltm",
            ".csv",
        }
        image_exts = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif", ".webp"}
        if ext == ".pdf":
            return input_path
        elif ext in word_exts:
            doc = None
            try:
                doc = word_app.Documents.Open(str(input_path), AddToRecentFiles=False)
                doc.ExportAsFixedFormat(
                    OutputFileName=str(output_path),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
                return output_path
            finally:
                if doc is not None:
                    doc.Close(False)
        elif ext in excel_exts:
            wb = None
            try:
                wb = excel_app.Workbooks.Open(str(input_path))
                wb.ExportAsFixedFormat(0, str(output_path))
                return output_path
            finally:
                if wb is not None:
                    wb.Close(False)
        elif ext in image_exts:
            img = Image.open(str(input_path))
            frames = []
            try:
                n_frames = getattr(img, "n_frames", 1)
                for i in range(n_frames):
                    try:
                        img.seek(i)
                    except EOFError:
                        break
                    frames.append(img.convert("RGB"))
                if not frames:
                    frames = [img.convert("RGB")]
                first_frame, rest_frames = frames[0], frames[1:]
                first_frame.save(
                    str(output_path),
                    "PDF",
                    resolution=100.0,
                    save_all=True,
                    append_images=rest_frames,
                )
                return output_path
            finally:
                try:
                    img.close()
                except:
                    pass
                for frame in frames:
                    try:
                        frame.close()
                    except:
                        pass
        return None

