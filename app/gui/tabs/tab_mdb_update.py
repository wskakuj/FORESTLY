"""
Forestly — Mixin: TabMdbUpdateMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from pathlib import Path
import threading
import traceback
import shutil
import pyodbc
import pythoncom

class TabMdbUpdateMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_mdb_update_tab(self, parent):
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
            card, text="Folder źródłowy z .mdb:", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.mdb_source_entry = ctk.CTkEntry(
            card, placeholder_text="Wskaż folder z oryginalnymi bazami", height=36
        )
        self.mdb_source_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.mdb_source_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))
        ctk.CTkLabel(
            card, text="Folder docelowy zapisu:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=(8, 15), sticky="w")
        self.mdb_output_entry = ctk.CTkEntry(
            card, placeholder_text="Gdzie zapisać poprawione bazy?", height=36
        )
        self.mdb_output_entry.grid(row=1, column=1, padx=5, pady=(8, 15), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda: self.select_dir(self.mdb_output_entry),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=(8, 15))
        self.mdb_start_btn = ctk.CTkButton(
            scroll_frame,
            text="Usuń 0 w bazach (MDB)",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=44,
            corner_radius=6,
            command=self.start_mdb_update_pipeline,
        )
        self.mdb_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_mdb_update_pipeline(self):
        source_folder = (
            self.mdb_source_entry.get().strip() if self.mdb_source_entry else ""
        )
        output_folder = (
            self.mdb_output_entry.get().strip() if self.mdb_output_entry else ""
        )
        if not source_folder or not Path(source_folder).exists():
            messagebox.showwarning(
                "Błąd", "Wybierz istniejący folder z oryginalnymi plikami .mdb."
            )
            return
        if not output_folder:
            messagebox.showwarning(
                "Błąd", "Wybierz folder docelowy dla poprawionych plików."
            )
            return
        if self.running:
            return
        self.last_output_dir = Path(output_folder)
        self._disable_ui_for_process()
        self.log(f"[USUWANIE 0 W MDB] Przetwarzanie plików z {source_folder}...")
        self.set_progress(0)
        threading.Thread(
            target=self.run_mdb_update_thread,
            args=(source_folder, output_folder),
            daemon=True,
        ).start()

    def run_mdb_update_thread(self, source_folder_str, output_folder_str):
        pythoncom.CoInitialize()
        try:
            self.update_status("Kopiowanie i modyfikacja baz .mdb", "#0078D7")
            src_dir = Path(source_folder_str)
            dst_dir = Path(output_folder_str)
            dst_dir.mkdir(parents=True, exist_ok=True)
            mdb_files = list(src_dir.glob("*.mdb"))
            if not mdb_files:
                raise Exception("Brak plików .mdb w wybranym folderze źródłowym.")
            total = len(mdb_files)
            self.start_progress_tracking(total, "Aktualizacja baz MDB")

            for idx, src_file in enumerate(mdb_files, start=1):
                self.check_stop()
                self.progress_current_file = src_file.name
                dst_file = dst_dir / src_file.name
                self.log(f"Przetwarzanie bazy: {src_file.name}")
                if src_file.resolve() != dst_file.resolve():
                    shutil.copy2(src_file, dst_file)
                conn_str = rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={dst_file};"
                try:
                    with pyodbc.connect(conn_str) as conn:
                        cursor = conn.cursor()
                        cursor.execute(
                            "SELECT DISTINCT ADRESS_FOREST FROM F_ARODES WHERE ADRESS_FOREST IS NOT NULL"
                        )
                        rows = cursor.fetchall()
                        update_count = 0
                        for row in rows:
                            old_adres = row.ADRESS_FOREST
                            if len(old_adres) > 13:
                                new_adres = old_adres[:11] + "  " + old_adres[13:]
                                if old_adres != new_adres:
                                    cursor.execute(
                                        "UPDATE F_ARODES SET ADRESS_FOREST = ? WHERE ADRESS_FOREST = ?",
                                        (new_adres, old_adres),
                                    )
                                    update_count += 1
                        conn.commit()
                        self.log(
                            f"  -> Zakończono bazę {src_file.name} (zmodyfikowano {update_count} rekordów)."
                        )
                except Exception as e:
                    self.log(
                        f"  -> Błąd podczas przetwarzania bazy {src_file.name}: {e}"
                    )
                self.set_progress(idx / total)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log("\nZAKOŃCZONO POMYŚLNIE EDYCJĘ BAZ MDB.")
            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Sukces", "Operacja na bazach .mdb zakończona pomyślnie."
                ),
            )
        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            pythoncom.CoUninitialize()
            self.running = False
            self.after(0, self.restore_all_buttons)

