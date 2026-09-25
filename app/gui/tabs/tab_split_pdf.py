"""
Forestly — Mixin: TabSplitPdfMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import traceback
import win32com.client
import pythoncom

from app.config import (
    is_file_locked,
)
from pypdf import PdfWriter

class TabSplitPdfMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_split_pdf_tab(self, parent):
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
        self.split_title_folder_entry = ctk.CTkEntry(
            card,
            placeholder_text="Folder ze stronami tytułowymi STRTYT*.docx",
            height=36,
        )
        self.split_title_folder_entry.grid(
            row=0, column=1, padx=5, pady=(15, 8), sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.split_title_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card, text="Folder z opisami:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=8, sticky="w")
        self.split_opisy_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder z plikami opisów", height=36
        )
        self.split_opisy_folder_entry.grid(row=1, column=1, padx=5, pady=8, sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.split_opisy_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=8)
        ctk.CTkLabel(
            card, text="Folder z raportami:", font=font_label, text_color="#E0E0E0"
        ).grid(row=2, column=0, padx=15, pady=8, sticky="w")
        self.split_raporty_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder z raportami Excel", height=36
        )
        self.split_raporty_folder_entry.grid(
            row=2, column=1, padx=5, pady=8, sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.split_raporty_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=2, column=2, padx=15, pady=8)
        ctk.CTkLabel(
            card, text="Folder docelowy:", font=font_label, text_color="#E0E0E0"
        ).grid(row=3, column=0, padx=15, pady=(8, 15), sticky="w")
        self.split_output_folder_entry = ctk.CTkEntry(
            card, placeholder_text="Folder wyjściowy dla rozdzielonych PDF", height=36
        )
        self.split_output_folder_entry.grid(
            row=3, column=1, padx=5, pady=(8, 15), sticky="ew"
        )
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.split_output_folder_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=3, column=2, padx=15, pady=(8, 8))
        self.split_scalony_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            card,
            text="Utwórz dodatkowo scalony PDF dla każdej wsi (w kolejności powstawania plików)",
            variable=self.split_scalony_var,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            fg_color="#0067C0", hover_color="#005A9E", checkbox_height=20, checkbox_width=20,
        ).grid(row=4, column=0, columnspan=3, padx=15, pady=(4, 15), sticky="w")
        self.split_pdf_btn = ctk.CTkButton(
            scroll_frame,
            text="Rozdziel na osobne PDF",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=44,
            corner_radius=6,
            command=self.start_split_pdf_pipeline,
        )
        self.split_pdf_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_split_pdf_pipeline(self):
        title_folder = (
            self.split_title_folder_entry.get().strip()
            if self.split_title_folder_entry
            else ""
        )
        opisy_folder = (
            self.split_opisy_folder_entry.get().strip()
            if self.split_opisy_folder_entry
            else ""
        )
        raporty_folder = (
            self.split_raporty_folder_entry.get().strip()
            if self.split_raporty_folder_entry
            else ""
        )
        output_folder = (
            self.split_output_folder_entry.get().strip()
            if self.split_output_folder_entry
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
        scalony = False
        var = getattr(self, "split_scalony_var", None)
        if var is not None:
            try:
                scalony = bool(var.get())
            except Exception:
                scalony = False
        self.last_output_dir = Path(output_folder)
        self._disable_ui_for_process()
        self.log(f"[ROZDZIELENIE PDF] Zapis do struktury drzewa"
                 + (" + scalone PDF wsi" if scalony else "") + "...")
        self.set_progress(0)
        threading.Thread(
            target=self.run_split_pdf_thread,
            args=(title_folder, opisy_folder, raporty_folder, output_folder, scalony),
            daemon=True,
        ).start()

    def run_split_pdf_thread(
            self, title_folder_str, opisy_folder_str, raporty_folder_str, output_folder_str,
            scalony=False,
    ):
        pythoncom.CoInitialize()
        word = None
        excel = None
        try:
            self.update_status("Rozdzielanie dokumentacji i generowanie PDF", "#0078D7")
            output_folder = Path(output_folder_str)
            output_folder.mkdir(parents=True, exist_ok=True)

            all_villages = self._get_all_villages(title_folder_str, opisy_folder_str, raporty_folder_str)
            if not all_villages:
                raise Exception("Nie odnaleziono plików z nazwami wsi w podanych folderach.")

            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            total = len(all_villages)
            created_dirs = 0
            self.start_progress_tracking(total, "Rozdzielanie PDF")

            for idx, village_name in enumerate(all_villages, start=1):
                if self.stop_event.is_set():
                    raise InterruptedError()

                self.progress_current_file = village_name

                self.log(f"Przetwarzanie wsi: {village_name}")
                village_out_dir = output_folder / village_name
                village_out_dir.mkdir(parents=True, exist_ok=True)

                try:
                    file_counter = 1
                    pdfy_wsi = []   # kolejność powstawania plików w folderze wsi

                    path_title = self.find_matching_file(Path(title_folder_str), village_name)
                    if path_title:
                        pdf_str = village_out_dir / f"{file_counter}_STR_TYT_{village_name}.pdf"
                        if is_file_locked(path_title):
                            self.log(f"  [Błąd] Plik tytułowy zablokowany: {path_title.name}")
                        else:
                            self.convert_office_to_pdf(path_title, pdf_str, word, excel)
                            pdfy_wsi.append(pdf_str)
                            file_counter += 1

                    path_opis = self.find_matching_file(Path(opisy_folder_str), village_name)
                    if path_opis:
                        if is_file_locked(path_opis):
                            self.log(f"  [Błąd] Plik opisu zablokowany: {path_opis.name}")
                        else:
                            pdf_opis = village_out_dir / f"{file_counter}_OPIS_{village_name}.pdf"
                            self.convert_office_to_pdf(path_opis, pdf_opis, word, excel)
                            pdfy_wsi.append(pdf_opis)
                            file_counter += 1

                    path_raport = self.find_matching_file(Path(raporty_folder_str), village_name)
                    if path_raport:
                        if is_file_locked(path_raport):
                            self.log(f"  [Błąd] Plik raportu zablokowany: {path_raport.name}")
                        else:
                            wb = None
                            try:
                                wb = excel.Workbooks.Open(str(path_raport))
                                for ws_idx in range(1, wb.Worksheets.Count + 1):
                                    if self.stop_event.is_set():
                                        raise InterruptedError()
                                    ws = wb.Worksheets(ws_idx)
                                    safe_ws_name = "".join(
                                        c for c in ws.Name if c.isalnum() or c in (" ", "_", "-")).strip()
                                    pdf_ws = village_out_dir / f"{file_counter}_RAPORT_{village_name}_{safe_ws_name}.pdf"
                                    ws.ExportAsFixedFormat(0, str(pdf_ws))
                                    pdfy_wsi.append(pdf_ws)
                                    file_counter += 1
                            except InterruptedError:
                                raise
                            except Exception as e:
                                self.log(f"  [Błąd] Problem z eksportem arkuszy: {e}")
                            finally:
                                if wb is not None:
                                    wb.Close(False)

                    # scalony PDF wsi — w tej samej kolejności, w jakiej pliki powstały
                    if scalony and len(pdfy_wsi) > 1:
                        try:
                            writer = PdfWriter()
                            for pdf_path in pdfy_wsi:
                                writer.append(str(pdf_path))
                            scalony_path = (village_out_dir /
                                            f"{village_name}_{self._nazwa_dokumentu()}.pdf")
                            with open(scalony_path, "wb") as f:
                                writer.write(f)
                            self.log(f"  [SCALONY] {scalony_path.name} "
                                      f"({len(pdfy_wsi)} plików PDF)")
                        except Exception as e:
                            self.log(f"  [Błąd scalania {village_name}]: {e}")

                    created_dirs += 1
                except InterruptedError:
                    raise
                except Exception as e:
                    self.log(f"Błąd przetwarzania: {e}")
                self.set_progress(idx / total)

            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log(f"\nZAKOŃCZONO POMYŚLNIE. Przetworzono foldery dla {created_dirs} wsi.")
            self.after(0, lambda: messagebox.showinfo("Sukces", "Rozdzielanie na PDF zakończone."))

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

