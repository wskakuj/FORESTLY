"""
Forestly — Mixin: TabExcelMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import traceback
import shutil
import win32com.client
import pythoncom

from app.config import (
    EXCEL_SHEET_DEFAULTS, is_file_locked,
)

class TabExcelMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_excel_tab(self, parent):
        # Konfiguracja głównej zakładki, aby rozciągała się na całe okno
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        # --- NOWE: Przewijalna ramka (ScrollableFrame) na całą zawartość ---
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)

        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)
        font_sheet = ctk.CTkFont(family="Segoe UI", size=12)

        # ZMIANA: Zamiast 'parent', główna karta jest przypinana do 'scroll_frame'
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
            card, text="Folder z plikami Excel:", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.excel_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Wskaż folder z plikami .xls / .xlsx", height=36
        )
        self.excel_folder_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.excel_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(
            card, text="Folder docelowy:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=8, sticky="w")
        self.excel_output_entry = ctk.CTkEntry(
            card,
            placeholder_text="Wskaż folder zapisu dla ułożonych plików Excel",
            height=36,
        )
        self.excel_output_entry.grid(row=1, column=1, padx=5, pady=8, sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.excel_output_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=8)

        self.include_subfolders_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            card,
            text="Przetwarzaj także podfoldery",
            variable=self.include_subfolders_var,
        ).grid(row=2, column=0, columnspan=3, padx=15, pady=(2, 10), sticky="w")

        fonts_frame = ctk.CTkFrame(card, fg_color="transparent")
        fonts_frame.grid(
            row=3, column=0, columnspan=3, padx=15, pady=(0, 10), sticky="ew"
        )
        fonts_frame.grid_columnconfigure(0, weight=1)
        fonts_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            fonts_frame,
            text="Dostosowanie rozmiaru czcionek w arkuszach:",
            font=font_label,
            text_color="#A0A0A0",
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")

        # === Globalne ustawienie czcionki ===
        global_frame = ctk.CTkFrame(
            fonts_frame, fg_color="#1E1E1E", border_width=1, border_color="#333333"
        )
        global_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        self.global_font_var = ctk.BooleanVar(value=False)
        self.global_font_cb = ctk.CTkCheckBox(
            global_frame,
            text="Zastosuj ten sam rozmiar do WSZYSTKICH arkuszy:",
            variable=self.global_font_var,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            text_color="#0078D7",
            command=self._toggle_global_font_size,
        )
        self.global_font_cb.pack(side="left", padx=10, pady=8)

        self.global_font_entry = ctk.CTkEntry(
            global_frame,
            width=70,
            height=30,
            state="disabled",
            fg_color="#2A2A2A",
            border_color="#333333",
            text_color="#666666",
        )
        self.global_font_entry.insert(0, "10")
        self.global_font_entry.pack(side="left", padx=(0, 10), pady=8)

        self.excel_font_entries = {}
        left_items = EXCEL_SHEET_DEFAULTS[::2]
        right_items = EXCEL_SHEET_DEFAULTS[1::2]
        total_rows = max(len(left_items), len(right_items))
        for idx in range(total_rows):
            if idx < len(left_items):
                sheet_name, start_row, font_size = left_items[idx]
                if sheet_name == "Sheet4":
                    continue
                left_row = ctk.CTkFrame(fonts_frame, fg_color="transparent")
                left_row.grid(row=idx + 2, column=0, padx=(0, 18), pady=6, sticky="ew")
                left_row.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    left_row,
                    text=f"{sheet_name} (od w. {start_row}):",
                    font=font_sheet,
                    text_color="#DDDDDD",
                    anchor="e",
                ).grid(row=0, column=0, padx=(0, 8), sticky="e")
                entry = ctk.CTkEntry(left_row, width=70, height=30)
                entry.insert(0, str(font_size))
                entry.grid(row=0, column=1, sticky="e")
                self.excel_font_entries[sheet_name] = {
                    "entry": entry,
                    "start_row": start_row,
                }
            if idx < len(right_items):
                sheet_name, start_row, font_size = right_items[idx]
                right_row = ctk.CTkFrame(fonts_frame, fg_color="transparent")
                right_row.grid(row=idx + 2, column=1, padx=(18, 0), pady=6, sticky="ew")
                right_row.grid_columnconfigure(0, weight=1)
                ctk.CTkLabel(
                    right_row,
                    text=f"{sheet_name} (od w. {start_row}):",
                    font=font_sheet,
                    text_color="#DDDDDD",
                    anchor="e",
                ).grid(row=0, column=0, padx=(0, 8), sticky="e")
                entry = ctk.CTkEntry(right_row, width=70, height=30)
                entry.insert(0, str(font_size))
                entry.grid(row=0, column=1, sticky="e")
                self.excel_font_entries[sheet_name] = {
                    "entry": entry,
                    "start_row": start_row,
                }
        if "REJ" in self.excel_font_entries and "Sheet4" not in self.excel_font_entries:
            self.excel_font_entries["Sheet4"] = self.excel_font_entries["REJ"]

        # === OPCJE USUWANIA KOLUMN ===
        delete_options_frame = ctk.CTkFrame(card, fg_color="transparent")
        delete_options_frame.grid(row=4, column=0, columnspan=3, padx=15, pady=(5, 15), sticky="ew")

        ctk.CTkLabel(
            delete_options_frame, text="Opcje usuwania kolumn (w arkuszach Sheet4 / REJ):", font=font_label,
            text_color="#A0A0A0"
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        self.remove_owners_var = ctk.BooleanVar(value=False)
        self.cb_remove_owners = ctk.CTkCheckBox(
            delete_options_frame, text="Usuń Właścicieli (Nazwisko, Adres, Współwłaśc., Udział)", variable=self.remove_owners_var,
            font=font_sheet, fg_color="#8B0000", hover_color="#A52A2A"
        )
        self.cb_remove_owners.grid(row=1, column=0, sticky="w", padx=(0, 20))

        self.remove_ls_var = ctk.BooleanVar(value=False)
        self.cb_remove_ls = ctk.CTkCheckBox(
            delete_options_frame, text="Usuń Użytek ew.", variable=self.remove_ls_var, font=font_sheet,
            fg_color="#8B0000", hover_color="#A52A2A"
        )
        self.cb_remove_ls.grid(row=1, column=1, sticky="w")

        # === PRZYCISK ===
        btn_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        btn_frame.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)

        self.excel_start_btn = ctk.CTkButton(
            btn_frame,
            text="Uruchom układanie Exceli",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=44,
            corner_radius=6,
            command=self.start_excel_pipeline,
        )
        self.excel_start_btn.grid(row=0, column=0, sticky="ew")

    def start_excel_pipeline(self):
        folder = (
            self.excel_folder_entry.get().strip() if self.excel_folder_entry else ""
        )
        output_folder = (
            self.excel_output_entry.get().strip() if self.excel_output_entry else ""
        )
        if not folder or not Path(folder).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z plikami Excel.")
            return
        if not output_folder:
            messagebox.showwarning(
                "Błąd", "Wybierz folder docelowy dla ułożonych kopii Exceli."
            )
            return
        if self.running:
            return
        font_config = {}

        # Sprawdź czy używamy globalnego ustawienia
        if self.global_font_var.get():
            # Globalne ustawienie - pobierz wartość i zastosuj do wszystkich arkuszy
            global_value = self.global_font_entry.get().strip()
            if not global_value.isdigit():
                messagebox.showwarning(
                    "Błąd", "Rozmiar czcionki w polu globalnym musi być liczbą."
                )
                return
            global_size = int(global_value)
            for sheet_name, data in self.excel_font_entries.items():
                font_config[sheet_name] = {
                    "start_row": data["start_row"],
                    "font_size": global_size,
                }
        else:
            # Indywidualne ustawienia - pobierz wartości z każdego pola
            for sheet_name, cfg in self.excel_font_entries.items():
                value = cfg["entry"].get().strip()
                if not value.isdigit():
                    messagebox.showwarning(
                        "Błąd",
                        f"Rozmiar czcionki dla arkusza '{sheet_name}' musi być liczbą.",
                    )
                    return
                font_config[sheet_name] = {
                    "start_row": cfg["start_row"],
                    "font_size": int(value),
                }

        self.last_output_dir = Path(output_folder)
        self._disable_ui_for_process()
        self.log(f"[EXCEL] URUCHOMIENIE PROCEDURY\nFolder: {folder}")
        self.set_progress(0)
        include_subfolders = (
                getattr(self, "include_subfolders_var", None)
                and self.include_subfolders_var.get()
        )
        # Pobierz stan checkboxów usuwania kolumn
        remove_owners = self.remove_owners_var.get()
        remove_ls = self.remove_ls_var.get()

        threading.Thread(
            target=self.run_excel_thread,
            args=(folder, output_folder, font_config, include_subfolders, remove_owners, remove_ls),
            daemon=True,
        ).start()

    def run_excel_thread(
            self, folder_str, output_folder_str, font_config, include_subfolders, remove_owners=False, remove_ls=False
    ):
        pythoncom.CoInitialize()
        excel = None
        _excel_pid = None
        try:
            folder = Path(folder_str)
            output_folder = Path(output_folder_str)
            files = (
                list(folder.rglob("*.xls*"))
                if include_subfolders
                else list(folder.glob("*.xls*"))
            )
            files = sorted(
                [f for f in files if f.is_file() and not f.name.startswith("~$")]
            )
            if not files:
                raise Exception("Brak plików Excel.")
            excel = win32com.client.DispatchEx("Excel.Application")
            try:
                from app.core import office_guard
                _excel_pid = office_guard.register(excel)
            except Exception:
                pass
            excel.Visible = False
            excel.DisplayAlerts = False
            total = len(files)
            self.start_progress_tracking(total, "Układanie Exceli")

            for idx, file_path in enumerate(files, start=1):
                self.check_stop()
                self.progress_current_file = file_path.name
                if is_file_locked(file_path):
                    self.log(f"POMINIĘTO (Plik zablokowany/otwarty): {file_path.name}")
                    continue
                self.log(f"Przetwarzanie: {file_path.name}")
                wb = None
                try:
                    rel_path = file_path.relative_to(folder)
                    target_path = output_folder / rel_path
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    if file_path.resolve() != target_path.resolve():
                        shutil.copy2(file_path, target_path)
                    wb = excel.Workbooks.Open(str(target_path))
                    self.process_excel_workbook(excel, wb, font_config)

                    # --- USUWANIE KOLUMN (jeśli zaznaczono) ---
                    if remove_owners or remove_ls:
                        # Nazwy nagłówków do wyszukania w wierszach 1-9
                        owner_headers = ["nazwisko i imię", "adres właściciela", "współwłaścicieli", "udział"]
                        ls_headers = ["użytek ew"]
                        for sheet_name in ["Sheet4", "REJ"]:
                            try:
                                ws = self.get_sheet_if_exists(wb, sheet_name)
                                if ws:
                                    cols_to_delete = set()
                                    # Przeszukaj wiersze 1-9, wszystkie kolumny
                                    for row in range(1, 10):
                                        for col in range(1, 51):
                                            cell_val = ws.Cells(row, col).Value
                                            if cell_val is not None:
                                                val_str = str(cell_val).strip().lower()
                                                if remove_owners:
                                                    for h in owner_headers:
                                                        if h in val_str:
                                                            cols_to_delete.add(col)
                                                            break
                                                if remove_ls:
                                                    for h in ls_headers:
                                                        if h in val_str:
                                                            cols_to_delete.add(col)
                                                            break
                                    # Usuń kolumny od końca (żeby nie popsuć indeksów)
                                    for col in sorted(cols_to_delete, reverse=True):
                                        ws.Columns(col).Delete()
                                        self.log(f"  -> Usunięto kolumnę {col} z arkusza {sheet_name}")
                            except Exception as e:
                                self.log(f"  [Ostrzeżenie] Problem z usuwaniem kolumn w {sheet_name}: {e}")
                    # -------------------------------------------

                    wb.Close(SaveChanges=True)
                except Exception as e:
                    self.log(f"Błąd pliku {file_path.name}: {e}")
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                self.set_progress(idx / total)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except:
                    pass
                try:
                    from app.core import office_guard
                    office_guard.unregister(_excel_pid)
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

    def process_excel_workbook(self, excel, wb, font_config):
        wb.CheckCompatibility = False
        self.reorder_sheets(wb)
        self.delete_unwanted_sheets(wb)
        self.setup_printing_and_styles(wb, excel)
        self.apply_font_sizes(wb, font_config)

        # --- NAPRAWA LITERÓWKI W REJESTR1 ---
        for sheet_name in ["Sheet4", "REJ"]:
            ws = None
            try:
                ws = wb.Sheets(sheet_name)
            except Exception:
                continue
            if ws is not None:
                try:
                    found = ws.Cells.Find("Wskazania godspodarcze")
                    while found:
                        found.Value = "Wskazania gospodarcze"
                        found = ws.Cells.FindNext(found)
                except Exception:
                    pass
        # -------------------------------------

        try:
            wb.Worksheets(1).Select()
        except Exception:
            pass

    def setup_printing_and_styles(self, wb, excel):
        xlLandscape = 2
        xlPortrait = 1
        xlPaperA4 = 9
        LIGHT_GRAY_COLOR = 0xE6E6E6
        for sheet_name in ["Zestawienie", "WykazPow", "Skroty"]:
            ws = self.get_sheet_if_exists(wb, sheet_name)
            if ws:
                try:
                    ps = ws.PageSetup
                    ps.Orientation = xlPortrait
                    ps.PaperSize = xlPaperA4
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                    ps.Zoom = False
                    ps.LeftMargin = excel.InchesToPoints(0.75)
                    ps.RightMargin = excel.InchesToPoints(0.6)
                    ps.TopMargin = excel.InchesToPoints(0.6)
                    ps.BottomMargin = excel.InchesToPoints(0.4)
                except Exception:
                    pass
        ws_ot = self.get_sheet_if_exists(wb, "OT")
        if ws_ot:
            try:
                rng_ot = ws_ot.Range("A5:U9")
                rng_ot.Interior.Color = LIGHT_GRAY_COLOR
                ps_ot = ws_ot.PageSetup
                ps_ot.PrintTitleRows = "$5:$9"
                ps_ot.Orientation = xlLandscape
                ps_ot.PaperSize = xlPaperA4
                ps_ot.FitToPagesWide = 1
                ps_ot.FitToPagesTall = False
                ps_ot.Zoom = False
                ps_ot.TopMargin = excel.CentimetersToPoints(2)
                ps_ot.BottomMargin = excel.CentimetersToPoints(1)
                ps_ot.LeftMargin = excel.CentimetersToPoints(0)
                ps_ot.RightMargin = excel.CentimetersToPoints(0)
            except Exception as e:
                self.log(f"  [Ostrzeżenie] Problem z formatowaniem OT: {e}")
        ws_rej = self.get_sheet_if_exists(wb, "REJ") or self.get_sheet_if_exists(
            wb, "Sheet4"
        )
        if ws_rej:
            try:
                rng_rej = ws_rej.Range("A5:Q9")
                rng_rej.Interior.Color = LIGHT_GRAY_COLOR
                ps_rej = ws_rej.PageSetup
                ps_rej.PrintTitleRows = "$5:$9"
                ps_rej.Orientation = xlLandscape
                ps_rej.PaperSize = xlPaperA4
                ps_rej.FitToPagesWide = 1
                ps_rej.FitToPagesTall = False
                ps_rej.Zoom = False
                ps_rej.TopMargin = excel.CentimetersToPoints(2)
                ps_rej.BottomMargin = excel.CentimetersToPoints(1)
                ps_rej.LeftMargin = excel.CentimetersToPoints(0)
                ps_rej.RightMargin = excel.CentimetersToPoints(0)
            except Exception as e:
                self.log(f"  [Ostrzeżenie] Problem z formatowaniem REJ/Sheet4: {e}")

    def reorder_sheets(self, wb):
        moves = [
            ("OT", 4),
            ("Zestawienie", 3),
            ("WykazPow", 4),
            ("WykazWlasc", 6),
            ("WykazDzialek", 8),
            ("Skroty", 9),
        ]
        for sheet_name, before_index in moves:
            try:
                ws = self.get_sheet_if_exists(wb, sheet_name)
                if ws and wb.Worksheets.Count >= before_index:
                    ws.Move(Before=wb.Worksheets(before_index))
            except:
                pass

    def delete_unwanted_sheets(self, wb):
        to_delete = [
            "WzUPUL",
            "WykazDoZal",
            "ZestLasNLas",
            "Hodowla",
            "Przedrebne",
            "OchrPrzyrody",
            "Etaty",
        ]
        for sheet_name in to_delete:
            try:
                ws = self.get_sheet_if_exists(wb, sheet_name)
                if ws:
                    ws.Delete()
            except:
                pass

    def get_sheet_if_exists(self, wb, name):
        try:
            return wb.Worksheets(name)
        except:
            return None

    def setup_printing(self, wb, excel):
        pass

    def apply_font_sizes(self, wb, font_config):
        for sheet_name, cfg in font_config.items():
            try:
                ws = self.get_sheet_if_exists(wb, sheet_name)
                if ws:
                    ws.Rows(f"{cfg['start_row']}:{ws.Rows.Count}").Font.Size = cfg[
                        "font_size"
                    ]
            except:
                pass

