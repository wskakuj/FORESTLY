"""
Forestly — Mixin: TabExcelZMdbMixin
"""

import customtkinter as ctk
from tkinter import messagebox
from tkinter import filedialog
from pathlib import Path
import threading
import traceback
import pandas as pd
import pyodbc

class TabExcelZMdbMixin:
    """Mixin dla ModernApp — metody zostały wyciągnięte z oryginalnego guipia.py."""
    pass

    def setup_excel_z_mdb_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll_frame.grid_columnconfigure(0, weight=1)
        font_label = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        font_btn = ctk.CTkFont(family="Segoe UI", size=13)

        card = ctk.CTkFrame(
            scroll_frame, fg_color="#252526", corner_radius=8,
            border_width=1, border_color="#333333",
        )
        card.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="new")
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            card, text="Plik źródłowy (.mdb):", font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(15, 8), sticky="w")
        self.excel_z_mdb_src_entry = ctk.CTkEntry(
            card, placeholder_text="Wskaż plik bazy danych MDB...", height=36
        )
        self.excel_z_mdb_src_entry.grid(row=0, column=1, padx=5, pady=(15, 8), sticky="ew")
        ctk.CTkButton(
            card, text="Wybierz Plik", image=self.icon_folder,
            command=lambda: self.select_file(self.excel_z_mdb_src_entry, [("Baza Access", "*.mdb")]),
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=0, column=2, padx=15, pady=(15, 8))

        ctk.CTkLabel(
            card, text="Zapisz Excel jako:", font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=(8, 15), sticky="w")
        self.excel_z_mdb_out_entry = ctk.CTkEntry(
            card, placeholder_text="Gdzie zapisać gotowy plik .xlsx?", height=36
        )
        self.excel_z_mdb_out_entry.grid(row=1, column=1, padx=5, pady=(8, 15), sticky="ew")

        # Funkcja pomocnicza do zapisu pliku xlsx
        def select_save_xlsx():
            f_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Skoroszyt Excel", "*.xlsx")],
                title="Zapisz plik ewidencji jako"
            )
            if f_path:
                self.excel_z_mdb_out_entry.delete(0, "end")
                self.excel_z_mdb_out_entry.insert(0, f_path)

        ctk.CTkButton(
            card, text="Przeglądaj", image=self.icon_folder,
            command=select_save_xlsx,
            width=110, height=36, font=font_btn, fg_color="#333333", hover_color="#444444",
        ).grid(row=1, column=2, padx=15, pady=(8, 15))

        self.excel_z_mdb_start_btn = ctk.CTkButton(
            scroll_frame, text="Wyciągnij dane z MDB", image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            fg_color="#0067C0", hover_color="#005A9E", height=44, corner_radius=6,
            command=self.start_excel_z_mdb_pipeline,
        )
        self.excel_z_mdb_start_btn.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="ew")

    def start_excel_z_mdb_pipeline(self):
        mdb_path = self.excel_z_mdb_src_entry.get().strip() if self.excel_z_mdb_src_entry else ""
        out_path = self.excel_z_mdb_out_entry.get().strip() if self.excel_z_mdb_out_entry else ""

        if not mdb_path or not Path(mdb_path).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący plik źródłowy .mdb.")
            return
        if not out_path:
            messagebox.showwarning("Błąd", "Wybierz miejsce i nazwę dla docelowego pliku Excel.")
            return

        if self.running:
            return

        self.last_output_dir = Path(out_path).parent
        self._disable_ui_for_process()
        self.log(f"[EXCEL Z MDB] URUCHOMIENIE\nZ: {mdb_path}\nDo: {out_path}")
        self.set_progress(0)

        threading.Thread(
            target=self.run_excel_z_mdb_thread, args=(mdb_path, out_path), daemon=True,
        ).start()

    def run_excel_z_mdb_thread(self, mdb_path_str, output_path_str):
        import warnings
        warnings.filterwarnings('ignore', category=UserWarning)
        try:
            self.update_status("Wyciąganie ewidencji z pliku MDB...", "#0078D7")
            self.start_progress_tracking(1, "Eksport bazy MDB")
            self.check_stop()

            self.log(f"  -> Nawiązywanie połączenia z bazą: {Path(mdb_path_str).name}")
            conn_str = rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path_str};"
            conn = pyodbc.connect(conn_str)

            sql = """
            SELECT 
                p.PARCEL_NR AS numer_dzialki, 
                p.LAND_REGISTER_NR AS j_rej, 
                p.PARCEL_AREA AS pow_dzialki, 
                u.AREA_USE_CD AS klasouzytek, 
                u.LAND_USE_AREA AS pow_klasouz
            FROM 
                F_PARCEL p
            LEFT JOIN 
                F_PARCEL_LAND_USE u ON p.PARCEL_INT_NUM = u.PARCEL_INT_NUM
            """

            self.log("  -> Pobieranie surowych danych i rozwiązywanie relacji...")
            df = pd.read_sql(sql, conn)
            conn.close()

            self.check_stop()
            self.log("  -> Formatowanie kolumn zgodnie ze standardami Kombajnu...")

            df = df.rename(columns={
                'numer_dzialki': 'Numer działki',
                'j_rej': 'J. rej.',
                'pow_dzialki': 'Pow. działki',
                'klasouzytek': 'Klasoużytek',
                'pow_klasouz': 'Pow. klasouż.'
            })

            df['Właściciel'] = 'Brak danych'
            df['Numer działki'] = df['Numer działki'].astype(str).str.strip()
            df['J. rej.'] = df['J. rej.'].astype(str).str.strip()

            df = df[['Numer działki', 'J. rej.', 'Pow. działki', 'Właściciel', 'Klasoużytek', 'Pow. klasouż.']]

            self.check_stop()
            self.log(f"  -> Trwa zapisywanie do pliku Excel: {Path(output_path_str).name}")
            df.to_excel(output_path_str, index=False)

            self.set_progress(1)
            self.update_status("Zakończono pomyślnie.", "#27ae60", animate=False)
            self.log(f"\n✅ SUKCES! Plik został wygenerowany. Zmodyfikuj ewentualne poprawki w kolumnie właścicieli i wczytaj do modułu Rozliczania powierzchni.")
            self.after(0, lambda: messagebox.showinfo("Sukces", "Wygenerowano plik Excel zgodny z systemem."))

        except InterruptedError:
            self.update_status("Przerwano", "#D83B01", animate=False)
            self.log("\nZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
        except Exception as e:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

