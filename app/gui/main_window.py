"""
Forestly — Główne okno aplikacji
===========================================
Zależności: config.py, models.py, updater.py, widgets/*, tabs/*
Odpowiada za: główną klasę ModernApp — inicjalizację UI, dashboard,
              historię, pasek postępu, live stream, oraz łączenie
              wszystkich mixinów (zakładek) w jedną klasę.

Architektura mixinów:
  ModernApp(TabAllMixin, TabWordMixin, ..., UpdaterMixin, ctk.CTk)
  Każdy mixin dostarcza metody setup_*_tab() i run_*_thread() dla danej zakładki.
"""

import os
import sys
import time
import json
import threading
import traceback
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageDraw

from app.config import (
    CURRENT_VERSION, COLORS, ENCODING, HISTORY_FILE, MARGINS_FILE, SETTINGS_FILE,
    TERRITORY_DATA, PDF_ORDER_TEMPLATES, EXCEL_SHEET_DEFAULTS,
    kill_orphan_office_processes, load_margins, save_margins,
    add_tooltip,
    get_saved_template_order, set_saved_template_order,
    build_ordered_pdfs_from_templates,
)
from app.models import PipelineStep, StreamFile
from app.updater import UpdaterMixin
from app.gui.widgets.pdf_order_window import PdfOrderWindow
from app.gui.widgets.manual_pdf_merge_window import ManualPdfMergeWindow
from app.gui.widgets.changelog_window import ChangelogWindow
from app.gui.widgets.validation_window import ValidationWindow

from app.gui.tabs.tab_all import TabAllMixin
from app.gui.tabs.tab_word import TabWordMixin
from app.gui.tabs.tab_pdf import TabPdfMixin
from app.gui.tabs.tab_manual_merge import TabManualMergeMixin
from app.gui.tabs.tab_template_generator import TabTemplateGeneratorMixin
from app.gui.tabs.tab_title_pages import TabTitlePagesMixin
from app.gui.tabs.tab_excel import TabExcelMixin
from app.gui.tabs.tab_layout_excel import TabLayoutExcelMixin
from app.gui.tabs.tab_split_pdf import TabSplitPdfMixin
from app.gui.tabs.tab_mdb_update import TabMdbUpdateMixin
from app.gui.tabs.tab_pdf_converter import TabPdfConverterMixin
from app.gui.tabs.tab_rozliczanie import TabRozliczanieMixin
from app.gui.tabs.tab_halizny import TabHaliznyMixin
from app.gui.tabs.tab_wydruki import TabWydrukiMixin
from app.gui.tabs.tab_excel_z_mdb import TabExcelZMdbMixin
from app.gui.tabs.tab_tworzenie_mietkow import TabTworzenieMietkowMixin
from app.gui.tabs.tab_nazwiska_mietek import TabNazwiskaMietekMixin
from app.gui.tabs.tab_mietek_rozbieznosci import TabMietekRozbieznosciMixin
from app.gui.tabs.tab_opis_og import TabOpisOgMixin


class ModernApp(
    TabAllMixin,
    TabWordMixin,
    TabPdfMixin,
    TabManualMergeMixin,
    TabTemplateGeneratorMixin,
    TabTitlePagesMixin,
    TabExcelMixin,
    TabLayoutExcelMixin,
    TabSplitPdfMixin,
    TabMdbUpdateMixin,
    TabPdfConverterMixin,
    TabRozliczanieMixin,
    TabHaliznyMixin,
    TabWydrukiMixin,
    TabExcelZMdbMixin,
    TabTworzenieMietkowMixin,
    TabNazwiskaMietekMixin,
    TabMietekRozbieznosciMixin,
    TabOpisOgMixin,
    UpdaterMixin,
    ctk.CTk,
):
    """Główna klasa aplikacji Kombajn Leśny PRO.

    Dziedziczy po wszystkich mixinach (jedna na zakładkę) oraz po ctk.CTk.
    Metody specyficzne dla zakładek znajdują się w app/gui/tabs/tab_*.py.
    Metody aktualizacji znajdują się w app/updater.py (UpdaterMixin).
    Ten plik zawiera tylko logikę wspólną: init, UI shell, dashboard, progress, stream.
    """

    def __init__(self):
        super().__init__()
        self.title("Forestly")

        # Dynamiczny rozmiar okna - 85% dostępnej przestrzeni ekranu
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)

        # Ograniczenia maksymalne i minimalne
        window_width = min(window_width, 1400)  # Maks 1400px
        window_height = min(window_height, 1000)  # Maks 1000px
        window_width = max(window_width, 900)  # Min 900px
        window_height = max(window_height, 650)  # Min 650px

        # Wycentrowanie okna
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.minsize(900, 650)  # Mniejszy minimalny rozmiar

        kill_orphan_office_processes()
        self.stop_event = threading.Event()
        self.running = False
        self.last_output_dir = None

        self._init_icons()
        self.entries = {}
        self.manual_pdf_src = None
        self.manual_pdf_dst = None
        self.excel_folder_entry = None
        self.excel_output_entry = None
        self.excel_start_btn = None
        self.excel_font_entries = {}
        self.tpl_data = {"MIETEK": {}, "TAKSATOR": {}}
        self.title_template_entry = None
        self.title_excel_entry = None
        self.title_output_entry = None
        self.title_village_placeholder_entry = None
        self.title_area_placeholder_entry = None
        self.title_generate_btn = None
        self.mietek_title_template_entry = None
        self.mietek_title_word_entry = None
        self.mietek_title_village_placeholder_entry = None
        self.mietek_title_area_placeholder_entry = None
        self.mietek_title_generate_btn = None
        self.layout_title_folder_entry = None
        self.layout_opisy_folder_entry = None
        self.layout_raporty_folder_entry = None
        self.layout_output_folder_entry = None
        self.layout_merge_btn = None
        self.pdfconv_source_entry = None
        self.pdfconv_output_entry = None
        self.pdfconv_start_btn = None
        self.split_title_folder_entry = None
        self.split_opisy_folder_entry = None
        self.split_raporty_folder_entry = None
        self.split_output_folder_entry = None
        self.split_pdf_btn = None
        self.mdb_source_entry = None
        self.mdb_output_entry = None
        self.mdb_start_btn = None
        self.rozl_xls_entry = None
        self.rozl_val_entry = None
        self.rozl_out_entry = None
        self.rozl_start_btn = None

        self.mietki_base_entry = None
        self.mietki_out_entry = None
        self.mietki_names_textbox = None
        self.mietki_start_btn = None
        self.nazwiska_bazowy_entry = None
        self.nazwiska_out_entry = None
        self.nazwiska_mietek_start_btn = None

        self.krzyz_xls_entry = None
        self.krzyz_mietki_entry = None
        self.krzyz_start_btn = None
        self.halizny_mietki_entry = None
        self.opis_og_root_entry = None
        self.opis_og_gdos_entry = None
        self.halizny_start_btn = None
        self.excel_z_mdb_src_entry = None
        self.excel_z_mdb_out_entry = None
        self.excel_z_mdb_start_btn = None

        # Zmienne dla Pełny Automat - SKROTY (STR_TYT powstaje z kreatora)
        self.all_gen_skroty_var = None

        self.all_skroty_entry = None
        self.all_skroty_frame = None

        self.status_base_text = "System oczekuje na zadanie"
        self.status_dots = 0

        # Live File Stream - śledzenie przetwarzania w czasie rzeczywistym
        self.stream_queue = []
        self.stream_current = None
        self.stream_completed = []
        self.stream_start_time = None
        self.stream_files_count = 0
        self.stream_frame = None
        self.stream_listbox = None
        self.stream_speed_label = None
        self.stream_eta_label = None

        self.build_ui()
        self.animate_status()
        self.check_pending_changelog()  # <--- SPRAWDŹ I WYŚWIETL CHANGELOG JEŚLI ISTNIEJE
        self.after(2000, lambda: self.check_github_update(manual=False))
    def check_pending_changelog(self):
        if getattr(sys, "frozen", False):
            app_dir = Path(sys.executable).resolve().parent
        else:
            app_dir = Path(__file__).resolve().parent

        changelog_file = app_dir / "pending_changelog.json"

        if changelog_file.exists():
            try:
                data = json.loads(changelog_file.read_text(encoding="utf-8"))
                version = data.get("version", CURRENT_VERSION)
                changelog_text = data.get("changelog", "")

                # Funkcja wywołująca okno dopiero po pełnym załadowaniu interfejsu
                def _show_window():
                    try:
                        win = ChangelogWindow(self, version, changelog_text)
                    except Exception as err:
                        print(f"[INFO] Błąd wyrysowania okna changelogu: {err}")

                # Wywołujemy po 1.5 sekundy od uruchomienia aplikacji
                self.after(1500, _show_window)

            except Exception as e:
                print(f"[INFO] Błąd odczytu pliku changelogu: {e}")
            finally:
                # Plik usuwamy dopiero po lekkim opóźnieniu, dając czas na jego przeczytanie
                def _safe_unlink():
                    try:
                        changelog_file.unlink(missing_ok=True)
                    except Exception:
                        pass

                self.after(3000, _safe_unlink)

    def load_history(self):
        if HISTORY_FILE.exists():
            try:
                return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            except Exception:
                return []
        return []

    def save_history(self, history):
        try:
            HISTORY_FILE.write_text(
                json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as e:
            print(f"[INFO] Błąd zapisu historii: {e}")

    def add_to_history(self, path):
        if not path or not Path(path).exists():
            return
        history = self.load_history()
        if path in history:
            history.remove(path)
        history.insert(0, path)
        history = history[:15]
        self.save_history(history)

    def show_history_menu(self, event, entry_widget):
        history = self.load_history()
        if not history:
            messagebox.showinfo("Historia", "Brak zapisanych folderów w historii.")
            return
        menu = tk.Menu(
            self,
            tearoff=0,
            bg="#252526",
            fg="#E0E0E0",
            activebackground="#0067C0",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 10),
            relief="flat",
            borderwidth=1,
        )
        for path in history:
            display_path = path if len(path) < 60 else "..." + path[-57:]
            menu.add_command(
                label=display_path,
                command=lambda p=path, e=entry_widget: self._insert_from_history(p, e),
            )
        menu.add_separator()
        menu.add_command(label="Wyczyść historię", command=self.clear_history)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _insert_from_history(self, path, entry_widget):
        entry_widget.delete(0, "end")
        entry_widget.insert(0, path)
        self.add_to_history(path)

    def load_settings(self):
        """Wczytuje zapisane ustawienia (pola WSIE.DBF, foldery, checkboxy)."""
        if SETTINGS_FILE.exists():
            try:
                return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_settings(self, settings):
        """Zapisuje ustawienia do pliku."""
        try:
            SETTINGS_FILE.write_text(
                json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception as e:
            print(f"[INFO] Błąd zapisu ustawień: {e}")

    def get_setting(self, key, default=""):
        """Pobiera pojedynczą wartość z ustawień."""
        return self.load_settings().get(key, default)

    def set_setting(self, key, value):
        """Ustawia pojedynczą wartość w ustawieniach."""
        settings = self.load_settings()
        settings[key] = value
        self.save_settings(settings)

    def clear_history(self):
        self.save_history([])
        messagebox.showinfo("Historia", "Historia folderów została wyczyszczona.")
    def build_dashboard_ui(self, parent):
        parent.grid_columnconfigure((0, 2, 4, 6, 8), weight=1)
        parent.grid_columnconfigure((1, 3, 5, 7), weight=0)
        self.dash_steps = []
        steps_info = [
            ("1. Halizny", "Generowanie"),
            ("2. TXT", "Czyszczenie"),
            ("3. Word", "Kompilacja"),
            ("4. PDF", "Konwersja"),
            ("5. Scalanie", "Integracja"),
        ]
        for i, (title, subtitle) in enumerate(steps_info):
            col = i * 2
            step_frame = ctk.CTkFrame(
                parent,
                fg_color="#252526",
                corner_radius=8,
                border_width=1,
                border_color="#333333",
            )
            step_frame.grid(row=0, column=col, padx=5, pady=10, sticky="ew")
            indicator = ctk.CTkLabel(
                step_frame, text="⚪", font=ctk.CTkFont(size=20), text_color="#555555"
            )
            indicator.pack(pady=(10, 0))
            lbl_title = ctk.CTkLabel(
                step_frame,
                text=title,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color="#E0E0E0",
            )
            lbl_title.pack()
            lbl_status = ctk.CTkLabel(
                step_frame,
                text=subtitle,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color="#888888",
            )
            lbl_status.pack(pady=(0, 10))
            self.dash_steps.append(
                {
                    "frame": step_frame,
                    "indicator": indicator,
                    "status": lbl_status,
                    "title": title,
                }
            )
            if i < len(steps_info) - 1:
                arrow = ctk.CTkLabel(
                    parent,
                    text="➔",
                    font=ctk.CTkFont(size=18, weight="bold"),
                    text_color="#555555",
                )
                arrow.grid(row=0, column=col + 1, padx=2, pady=5)

    def update_dashboard(self, step_index, status, text=None):
        if not hasattr(self, "dash_steps") or step_index >= len(self.dash_steps):
            return
        step = self.dash_steps[step_index]

        def _update():
            if status == "pending":
                step["indicator"].configure(text="⚪", text_color="#555555")
                step["frame"].configure(border_color="#333333")
                if text:
                    step["status"].configure(text=text, text_color="#888888")
            elif status == "running":
                step["indicator"].configure(text="🔄", text_color="#0078D7")
                step["frame"].configure(border_color="#0078D7")
                if text:
                    step["status"].configure(text=text, text_color="#0078D7")
            elif status == "done":
                step["indicator"].configure(text="✅", text_color="#27ae60")
                step["frame"].configure(border_color="#27ae60")
                if text:
                    step["status"].configure(text=text, text_color="#27ae60")
            elif status == "error":
                step["indicator"].configure(text="❌", text_color="#D83B01")
                step["frame"].configure(border_color="#D83B01")
                if text:
                    step["status"].configure(text=text, text_color="#D83B01")

        self.after(0, _update)

    def reset_dashboard(self):
        if hasattr(self, "dash_steps"):
            subtitles = [
                "Generowanie",
                "Czyszczenie",
                "Kompilacja",
                "Konwersja",
                "Integracja",
            ]
            for i, step in enumerate(self.dash_steps):
                self.update_dashboard(i, "pending", subtitles[i])

    def _create_fallback_icon(self, color, shape_type="circle"):
        img = Image.new("RGBA", (20, 20), color=(0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if shape_type == "circle":
            d.ellipse((2, 2, 18, 18), fill=color)
        elif shape_type == "play":
            d.polygon([(6, 4), (6, 16), (16, 10)], fill=color)
        elif shape_type == "stop":
            d.rectangle((5, 5, 15, 15), fill=color)
        return ctk.CTkImage(light_image=img, dark_image=img, size=(16, 16))

    def _init_icons(self):
        self.icon_folder = self._create_fallback_icon("#FFD700", "circle")
        self.icon_start = self._create_fallback_icon("#00FA9A", "play")
        self.icon_stop = self._create_fallback_icon("#DC143C", "stop")
    def open_last_output_dir(self):
        if self.last_output_dir and Path(self.last_output_dir).exists():
            # finalne wyniki 1-Click: od razu folder 'PDF polaczone'
            target = Path(self.last_output_dir)
            final = target / "PDF polaczone"
            if final.exists():
                target = final
            os.startfile(str(target))
        else:
            messagebox.showwarning(
                "Brak folderu",
                "Nie odnaleziono folderu docelowego lub żaden proces nie został jeszcze wykonany.",
            )
    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)
        left_header = ctk.CTkFrame(header_frame, fg_color="transparent")
        left_header.pack(side="left")
        ctk.CTkLabel(
            left_header,
            text="KOMBAJN LEŚNY",
            font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"),
            text_color="#E0E0E0",
        ).pack(side="left")
        ctk.CTkLabel(
            left_header,
            text="PRO",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#0078D7",
        ).pack(side="left", padx=(6, 0), pady=(8, 0))
        ctk.CTkLabel(
            left_header,
            text="System automatyzacji dokumentacji",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color="#888888",
        ).pack(side="left", padx=(20, 0), pady=(8, 0))

        self.btn_update = ctk.CTkButton(
            header_frame,
            text=f"Wersja {CURRENT_VERSION} (Sprawdź update)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="transparent",
            border_width=1,
            border_color="#0078D7",
            text_color="#0078D7",
            hover_color="#1E1E1E",
            height=28,
            command=lambda: self.check_github_update(manual=True),
        )
        self.btn_update.pack(side="right")
        add_tooltip(
            self.btn_update,
            "Połącz z GitHubem i sprawdź, czy dostępna jest nowsza wersja programu.",
        )

        self.top_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.top_panel.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="nsew")
        self.top_panel.grid_columnconfigure(0, weight=1)
        self.top_panel.grid_rowconfigure(1, weight=1)

        self.bottom_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_panel.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.bottom_panel.grid_columnconfigure(0, weight=1)

        # --- PIONOWE MENU (zamiast zakładek u góry) ---
        self.aktualna_zakladka = None          # klucz "SEKCJA|Nazwa"
        self.aktualna_zakladka_nazwa = None
        self._zakladka_frames = {}
        self._zakladka_buttons = {}

        self.top_panel.grid_columnconfigure(0, weight=0)
        self.top_panel.grid_columnconfigure(1, weight=1)

        sidebar = ctk.CTkScrollableFrame(
            self.top_panel, width=300, corner_radius=6,
            fg_color=("gray88", "gray17"),
        )
        sidebar.grid(row=1, column=0, sticky="nsw", padx=(0, 8), pady=8)
        self.sidebar = sidebar

        self.content_area = ctk.CTkFrame(self.top_panel, fg_color="transparent")
        self.content_area.grid(row=1, column=1, sticky="nsew")
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        font_btn_menu = ctk.CTkFont(family="Segoe UI", size=13)
        font_naglowka = ctk.CTkFont(family="Segoe UI", size=12, weight="bold")

        # zwijane grupy sekcji — domyślnie zwinięte (klik w nagłówek rozwija)
        self._sekcje_przyciski = {}
        self._sekcje_odstepy = {}
        self._sekcje_zwiniete = set()
        _aktualna_sekcja = [None]

        def _naglowek(tekst):
            sekcja = tekst
            przyciski = []
            self._sekcje_przyciski[sekcja] = przyciski
            _aktualna_sekcja[0] = sekcja

            def _przelacz(s=sekcja, lst=przyciski):
                kotwica = self._sekcje_odstepy.get(s)
                if s in self._sekcje_zwiniete:
                    self._sekcje_zwiniete.discard(s)
                    for b in lst:
                        if kotwica is not None:
                            b.pack(fill="x", padx=6, pady=2, before=kotwica)
                        else:
                            b.pack(fill="x", padx=6, pady=2)
                else:
                    self._sekcje_zwiniete.add(s)
                    for b in lst:
                        b.pack_forget()
                _odswiez(s)

            def _odswiez(s=sekcja):
                naglowek_btn.configure(
                    text=("▸ " if s in self._sekcje_zwiniete else "▾ ") + s)

            naglowek_btn = ctk.CTkButton(
                sidebar, text="▸ " + sekcja, font=font_naglowka, anchor="w",
                fg_color="transparent", hover_color=("gray75", "gray28"),
                text_color=("gray35", "gray60"), corner_radius=4,
                command=_przelacz,
            )
            naglowek_btn.pack(fill="x", padx=8, pady=(14, 2))
            self._sekcje_zwiniete.add(sekcja)  # startowo zwinięte
            # niewidoczny odstęp-zakładka: przyciski sekcji pakują się przed
            # nim, więc rozwinięcie grupy NIE spada na koniec menu
            odstep = ctk.CTkFrame(sidebar, fg_color="transparent", height=0)
            odstep.pack(fill="x")
            self._sekcje_odstepy[sekcja] = odstep

        def _nowa_zakladka(sekcja, nazwa, tooltip=None):
            klucz = f"{sekcja}|{nazwa}"
            ramka = ctk.CTkFrame(self.content_area, fg_color="transparent")
            ramka.grid(row=0, column=0, sticky="nsew")
            ramka.grid_columnconfigure(0, weight=1)
            ramka.grid_rowconfigure(0, weight=1)
            ramka.grid_remove()
            self._zakladka_frames[klucz] = ramka
            btn = ctk.CTkButton(
                sidebar, text=nazwa, anchor="w", height=34, corner_radius=6,
                font=font_btn_menu, fg_color="transparent",
                hover_color=("gray75", "gray28"),
                text_color=("gray10", "gray90"),
                command=lambda k=klucz: self.pokaz_zakladke(k),
            )
            kotwica = self._sekcje_odstepy.get(sekcja)
            if kotwica is not None:
                btn.pack(fill="x", padx=6, pady=2, before=kotwica)
            else:
                btn.pack(fill="x", padx=6, pady=2)
            self._zakladka_buttons[klucz] = btn
            przyciski_sekcji = self._sekcje_przyciski.get(sekcja)
            if przyciski_sekcji is not None:
                przyciski_sekcji.append(btn)
                if sekcja in self._sekcje_zwiniete:
                    btn.pack_forget()  # grupa startowo zwinięta
            if tooltip:
                add_tooltip(btn, tooltip)
            return ramka

        # ---- MIETEK ----
        _naglowek("MIETEK")
        tab_all = _nowa_zakladka(
            "MIETEK", "Pełny Automat (1-Click)",
            "Kompleksowy proces: czyści TXT, generuje i układa Worda, "
            "konwertuje na PDF i scala w gotowy dokument.")
        tab_wydruki_mietek = _nowa_zakladka("MIETEK", "Generowanie: MIETEK -> TXT")
        tab_word = _nowa_zakladka(
            "MIETEK", "Konwersja: MIETEK -> Word",
            "Tylko etap 1: Oczyszcza surowe pliki z systemu MIETEK i układa pliki Word.")
        tab_pdf = _nowa_zakladka(
            "MIETEK", "Konwersja: Word -> PDF",
            "Tylko etap 2: Zamienia gotowe pliki word na PDF i łączy w jeden plik.")
        tab_mietek_tpl_gen = _nowa_zakladka(
            "MIETEK", "Kreator Stron tytułowych",
            "Generuje jeden bazowy dokument Word ze stroną tytułową "
            "na podstawie wpisanych danych.")
        tab_mietek_title = _nowa_zakladka(
            "MIETEK", "Tworzenie Stron tytułowych",
            "Masowo tworzy strony tytułowe dla każdej wsi (MIETEK), "
            "wciągając dane z plików Word (OPTAX).")
        tab_opis_og = _nowa_zakladka(
            "MIETEK", "Opisy og\u00f3lne",
            "Tworzy \u201eopis og_<wie\u015b>.docx\u201d w folderach wsi \u2014 z Worda (WSK_ZB.doc) lub z danych MIETEKA (tymczasowo).")
        tab_mietek_rozb = _nowa_zakladka("MIETEK", "Wykaz Rozbieżności")
        tab_nazwiska_mietek = _nowa_zakladka(
            "MIETEK", "NAZWISKA -> MIETEK",
            "Klonuje strukturę MS-DOS i generuje W*.DBF pobierając nazwiska "
            "wyłącznie na podstawie pliku Ewidencji XLS.")
        tab_gdos_baza = _nowa_zakladka(
            "MIETEK", "Baza obszarów GDOŚ",
            "Edytor bazy obszarów ochrony przyrody (gdos_obszary.json) — kody, "
            "publikacje PZO i powiązania wstawiane do opisów ogólnych; z importem "
            "i eksportem do Excela.")
        tab_manual = _nowa_zakladka(
            "MIETEK", "Ręczne scalanie PDF",
            "Moduł ręczny: pozwala wczytać luźne PDF-y, poukładać je myszką "
            "w odpowiedniej kolejności i połączyć.")


        # ---- TAKSATOR ----
        _naglowek("TAKSATOR")
        tab_excel = _nowa_zakladka(
            "TAKSATOR", "Układanie Exceli do druku",
            "Optymalizuje pliki Excel: ukrywa zbędne arkusze, sortuje je "
            "i dostosowuje wielkość czcionki do druku.")
        tab_layout_excel = _nowa_zakladka(
            "TAKSATOR", "Wyłożenie Exceli",
            "Pobiera strony tytułowe, opisy i raporty, a następnie scala je "
            "w gotowe, pełne paczki PDF dla każdej wsi.")
        tab_template_gen = _nowa_zakladka(
            "TAKSATOR", "Kreator Stron tytułowych",
            "Generuje jeden bazowy dokument Word ze stroną tytułową "
            "na podstawie wpisanych danych.")
        tab_title = _nowa_zakladka(
            "TAKSATOR", "Tworzenie Stron tytułowych",
            "Masowo tworzy strony tytułowe dla każdej wsi, wciągając dane z zestawień Excel.")
        tab_opis_og_taks = _nowa_zakladka(
            "TAKSATOR", "Opisy ogólne",
            "Tworzy opisy ogólne (opis og_<wieś>.docx) na podstawie raportów Excel do druku.")
        tab_split_pdf = _nowa_zakladka(
            "TAKSATOR", "Excel -> PDF",
            "Konwertuje raporty i opisy, zachowując je jako osobne pliki PDF "
            "podzielone na foldery dla poszczególnych wsi.")
        tab_mdb_update = _nowa_zakladka(
            "TAKSATOR", "Usuwanie 0 w MDB",
            "Kopiuje bazy Access (.mdb) do nowego folderu i modyfikuje "
            "adresy leśne w tabeli F_ARODES.")

        # ---- ROZLICZANIE ----
        _naglowek("ROZLICZANIE")
        tab_rozl_main = _nowa_zakladka("ROZLICZANIE", "Rozliczanie powierzchni")
        tab_tworzenie_mietkow = _nowa_zakladka("ROZLICZANIE", "Tworzenie i wpisywanie mietków")
        tab_halizny = _nowa_zakladka("ROZLICZANIE", "Halizny")
        tab_zestawienie = _nowa_zakladka(
            "ROZLICZANIE", "Zestawienie zbiorcze",
            "Składa wszystkie pliki <WIEŚ>_Rozliczone.xlsx w jeden plik: "
            "sumy per wieś + rozpiska działek przybyło/ubyło z właścicielami."
        )
        tab_excel_z_mdb = _nowa_zakladka("ROZLICZANIE", "Excel z MDB")

        # ---- KONWERTER PDF ----
        _naglowek("KONWERTER PDF")
        tab_pdfconv = _nowa_zakladka("Konwerter PDF", "Konwerter PDF")

        # ZMIENIONE: Dodano extra_ui_setup=self._setup_all_extras oraz dashboard=True dla zakładki ALL
        self.setup_tab(
            tab_all,
            "ALL",
            "Folder źródłowy (MIETEK):",
            "Folder docelowy (struktura etapów):",
            show_order_button=True,
            extra_ui_setup=self._setup_all_extras,
            dashboard=True,
        )
        self.setup_wydruki_tab(tab_wydruki_mietek)
        self.setup_tab(
            tab_word,
            "WORD",
            "Folder źródłowy (MIETEK):",
            "Folder docelowy (TXT oraz Word):",
            show_order_button=False,
            extra_ui_setup=self._setup_word_extras,
        )
        self.setup_template_generator_tab(tab_mietek_tpl_gen, "MIETEK")
        self.setup_mietek_title_pages_tab(tab_mietek_title)
        self.setup_tab(
            tab_pdf,
            "PDF",
            "Folder źródłowy (Word):",
            "Folder docelowy (PDF):",
            show_order_button=True,
            extra_ui_setup=self._setup_pdf_extras,
        )
        self.setup_manual_merge_tab(tab_manual)
        self.setup_mietek_rozbieznosci_tab(tab_mietek_rozb)
        self.setup_nazwiska_mietek_tab(tab_nazwiska_mietek)
        self.setup_opis_og_tab(tab_opis_og)
        self.setup_gdos_editor_tab(tab_gdos_baza)

        self.setup_template_generator_tab(tab_template_gen, "TAKSATOR")
        self.setup_title_pages_tab(tab_title)
        self.setup_opis_og_taksator_tab(tab_opis_og_taks)
        self.setup_excel_tab(tab_excel)
        self.setup_layout_excel_tab(tab_layout_excel)
        self.setup_split_pdf_tab(tab_split_pdf)
        self.setup_mdb_update_tab(tab_mdb_update)

        self.setup_rozliczanie_tab(tab_rozl_main)
        self.setup_zestawienie_tab(tab_zestawienie)
        self.setup_tworzenie_mietkow_tab(tab_tworzenie_mietkow)
        self.setup_halizny_tab(tab_halizny)
        self.setup_excel_z_mdb_tab(tab_excel_z_mdb)
        self.setup_pdf_converter_tab(tab_pdfconv)

        self.options_frame = ctk.CTkFrame(self.top_panel, fg_color="transparent")
        self.options_frame.grid(row=2, column=1, pady=(5, 5), sticky="w")
        self.remove_names_var = ctk.BooleanVar(value=True)
        self.cb_remove_names = ctk.CTkCheckBox(
            self.options_frame,
            text="Usuwaj nazwiska z REJESTRU (oraz 1. stronę)",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            variable=self.remove_names_var,
            fg_color="#0067C0",
            hover_color="#005A9E",
        )
        self.cb_remove_names.pack(side="left", padx=5)
        add_tooltip(
            self.cb_remove_names,
            "Włączenie tej opcji uruchamia makra 'ZamienLF' oraz 'UsunNazwiskaRej', a także kasuje pierwszą stronę z rejestru.",
        )

        # startowo: pierwsza zakładka sekcji MIETEK
        self.pokaz_zakladke("MIETEK|Pełny Automat (1-Click)")

        log_frame = ctk.CTkFrame(self.bottom_panel, corner_radius=6)
        log_frame.grid(row=0, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        status_bar = ctk.CTkFrame(log_frame, fg_color="transparent")
        status_bar.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 0))
        status_bar.grid_columnconfigure(0, weight=1)
        self.status_label = ctk.CTkLabel(
            status_bar,
            text="Gotowy",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#0078D7",
        )
        self.status_label.grid(row=0, column=0, sticky="w")
        right_status_frame = ctk.CTkFrame(status_bar, fg_color="transparent")
        right_status_frame.grid(row=0, column=1, sticky="e")
        self.open_dir_btn = ctk.CTkButton(
            right_status_frame,
            text="Otwórz folder docelowy",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#27ae60",
            hover_color="#219653",
            height=28,
            image=self.icon_folder,
            command=self.open_last_output_dir,
            state="disabled",
        )
        self.open_dir_btn.pack(side="left", padx=(0, 8))
        add_tooltip(
            self.open_dir_btn,
            "Otwiera w Eksploratorze Windows folder z wygenerowanymi plikami.",
        )
        self.stop_btn = ctk.CTkButton(
            right_status_frame,
            text="Przerwij zadanie",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#8B0000",
            hover_color="#A52A2A",
            height=28,
            width=120,
            command=self.cancel_process,
            image=self.icon_stop,
            state="disabled",
        )
        self.stop_btn.pack(side="left")
        add_tooltip(self.stop_btn, "Bezpiecznie przerywa działanie obecnego zadania.")

        self.textbox = ctk.CTkTextbox(
            log_frame,
            height=140,
            font=ctk.CTkFont(family="Consolas", size=12),
            bg_color="transparent",
            fg_color="#1E1E1E",
            text_color="#D4D4D4",
            border_width=1,
            border_color="#333333",
        )
        self.textbox.grid(row=1, column=0, padx=15, pady=(8, 15), sticky="nsew")
        self.textbox.insert(
            "0.0", "System aktywny. Skonfiguruj proces i rozpocznij działanie.\n"
        )
        self.textbox.configure(state="disabled")
        # === NOWY PASEK POSTĘPU ZE SZCZEGÓŁAMI ===
        progress_container = ctk.CTkFrame(log_frame, fg_color="transparent")
        progress_container.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="ew")
        progress_container.grid_columnconfigure(0, weight=1)

        progress_info = ctk.CTkFrame(progress_container, fg_color="transparent")
        progress_info.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        progress_info.grid_columnconfigure(1, weight=1)

        self.progress_percent_label = ctk.CTkLabel(
            progress_info,
            text="0%",
            width=45,
            anchor="w",
            font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            text_color="#0078D7",
        )
        self.progress_percent_label.grid(row=0, column=0, sticky="w")

        self.progress_detail_label = ctk.CTkLabel(
            progress_info,
            text="Oczekiwanie na zadanie",
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#A0A0A0",
        )
        self.progress_detail_label.grid(row=0, column=1, sticky="w", padx=(8, 8))

        self.progress_eta_label = ctk.CTkLabel(
            progress_info,
            text="",
            width=160,
            anchor="e",
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#888888",
        )
        self.progress_eta_label.grid(row=0, column=2, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(
            progress_container, mode="determinate", height=8, progress_color="#0078D7"
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew")
        self.progress_bar.set(0)

        self.progress_total = 0
        self.progress_current = 0
        self.progress_start_time = None
        self.progress_current_file = None
        self.progress_description = ""

        # === LIVE FILE STREAM PANEL ===
        self.bottom_panel.grid_rowconfigure(0, weight=1)
        self.bottom_panel.grid_rowconfigure(1, weight=1)

        stream_panel = ctk.CTkFrame(self.bottom_panel, corner_radius=6)
        stream_panel.grid(
            row=1, column=0, sticky="nsew", pady=(5, 0)
        )  # było pady=(10, 0)
        stream_panel.grid_columnconfigure(0, weight=1)

        stream_header = ctk.CTkFrame(stream_panel, fg_color="transparent")
        stream_header.grid(
            row=0, column=0, sticky="ew", padx=12, pady=(8, 0)
        )  # mniejsze paddingi
        stream_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            stream_header,
            text="📊 Przetwarzanie w czasie rzeczywistym",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#0078D7",
        )  # było size=14

        stats_frame = ctk.CTkFrame(stream_header, fg_color="transparent")
        stats_frame.grid(row=0, column=1, sticky="e")

        self.stream_speed_label = ctk.CTkLabel(
            stats_frame,
            text="Prędkość: 0.0 plików/s",
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#888888",
        )
        self.stream_speed_label.pack(side="left", padx=(0, 20))

        self.stream_eta_label = ctk.CTkLabel(
            stats_frame,
            text="ETA: obliczanie...",
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="#888888",
        )
        self.stream_eta_label.pack(side="left")

        self.stream_listbox = tk.Listbox(
            stream_panel,
            font=("Consolas", 9),
            activestyle="none",  # było size=10
            selectmode=tk.SINGLE,
            bg="#1E1E1E",
            fg="#D4D4D4",
            selectbackground="#005A9E",
            borderwidth=0,
            highlightthickness=0,
            height=5,
        )  # było height=6
        self.stream_listbox.grid(
            row=1, column=0, sticky="nsew", padx=12, pady=(6, 10)
        )  # mniejsze paddingi

        stream_panel.grid_remove()  # Ukryty na starcie
        self.stream_frame = stream_panel

        self.update_options_visibility()

    def pokaz_zakladke(self, klucz):
        """Przełącza widok na zakładkę o kluczu 'SEKCJA|Nazwa' (pionowe menu)."""
        for k, ramka in self._zakladka_frames.items():
            if k == klucz:
                ramka.grid()
            else:
                ramka.grid_remove()
        for k, btn in self._zakladka_buttons.items():
            if k == klucz:
                btn.configure(fg_color="#0067C0", text_color="#FFFFFF")
            else:
                btn.configure(fg_color="transparent", text_color=("gray10", "gray90"))
        self.aktualna_zakladka = klucz
        self.aktualna_zakladka_nazwa = klucz.split("|", 1)[1] if klucz else None
        self.update_options_visibility()
        # wejście w Pełny Automat (1-Click) otwiera kreatora — jak w webowym GUI
        if klucz == "MIETEK|Pełny Automat (1-Click)":
            self.after(150, self._open_all_wizard)

    def update_options_visibility(self):
        of = getattr(self, "options_frame", None)
        if of is None:
            return
        nazwa = getattr(self, "aktualna_zakladka_nazwa", None)
        if nazwa == "Konwersja: MIETEK -> Word":
            of.grid()
        else:
            of.grid_remove()

    def update_status(self, text, color="#0078D7", animate=True):
        self.status_color = color
        def _update_stat():
            self.status_base_text = text
            self.status_label.configure(text_color=color)
            if not animate:
                self.status_label.configure(text=text)

        self.after(0, _update_stat)

    def animate_status(self):
        if self.running:
            dots = "." * (self.status_dots % 4)
            self.status_label.configure(text=f"{self.status_base_text}{dots}")
            self.status_dots += 1
            self.after(500, self.animate_status)

    def cancel_process(self):
        if self.running:
            self.log(
                "\n[SYSTEM] Wydano polecenie zatrzymania. Trwa awaryjne przerywanie procesów..."
            )
            self.stop_event.set()
            self.stop_btn.configure(state="disabled", text="Zatrzymywanie...")
            # po chwili ubij procesy Office zostawione w tle — zwalniają
            # blokady na folderach wynikowych (patrz app/core/office_guard)
            threading.Thread(target=self._office_sweep_after_stop,
                             daemon=True).start()

    def _office_sweep_after_stop(self):
        time.sleep(2.5)
        try:
            from app.core import office_guard
            office_guard.kill_registered(log=self.log)
        except Exception:
            pass

    # === METODY LIVE FILE STREAM ===
    def init_live_stream(self, total_files=None):
        self.stream_queue = []
        self.stream_current = None
        self.stream_completed = []
        self.stream_start_time = time.time()
        self.stream_files_count = 0
        if self.stream_listbox:
            self.stream_listbox.delete(0, tk.END)
        if self.stream_speed_label:
            self.stream_speed_label.configure(text="Prędkość: 0.0 plików/s")
        if self.stream_eta_label:
            self.stream_eta_label.configure(text="ETA: obliczanie...")

    def add_to_stream_queue(self, source_path, target_path=None):
        self.stream_queue.append(
            {
                "source": str(source_path),
                "target": str(target_path) if target_path else "...",
                "status": "pending",
            }
        )
        self.update_stream_display()

    def start_stream_file(self, source_path, target_path=None):
        self.stream_current = {
            "source": str(source_path),
            "target": str(target_path) if target_path else "...",
            "start_time": time.time(),
            "status": "processing",
        }
        # Usuń z kolejki, jeśli tam był
        self.stream_queue = [
            q for q in self.stream_queue if q["source"] != str(source_path)
        ]
        self.update_stream_display()

    def complete_stream_file(self, source_path, target_path, duration=None):
        if duration is None and self.stream_current:
            duration = time.time() - self.stream_current["start_time"]
        self.stream_completed.append(
            {
                "source": str(source_path),
                "target": str(target_path),
                "duration": duration,
                "status": "completed",
            }
        )
        self.stream_files_count += 1
        self.stream_current = None
        if len(self.stream_completed) > 15:
            self.stream_completed = self.stream_completed[-15:]
        self.update_stream_display()

    def update_stream_display(self):
        if not self.stream_listbox:
            return

        def _update():
            self.stream_listbox.delete(0, tk.END)
            recent_completed = (
                self.stream_completed[-5:]
                if len(self.stream_completed) > 5
                else self.stream_completed
            )
            for item in recent_completed:
                source_name = Path(item["source"]).name
                target_name = Path(item["target"]).name

                # Wymuszenie typu ułamkowego przed formatowaniem
                try:
                    duration_val = float(item["duration"])
                    duration_str = f"{duration_val:.1f}s"
                except (TypeError, ValueError):
                    duration_str = ""

                self.stream_listbox.insert(
                    tk.END, f"✅ {source_name} → {target_name} ({duration_str})"
                )
            if self.stream_current:
                source_name = Path(self.stream_current["source"]).name
                target_name = Path(self.stream_current["target"]).name
                self.stream_listbox.insert(
                    tk.END, f"🔄 {source_name} → {target_name}..."
                )
            pending = self.stream_queue[:3] if self.stream_queue else []
            for item in pending:
                source_name = Path(item["source"]).name
                self.stream_listbox.insert(tk.END, f"⏳ {source_name} (w kolejce)")
            if self.stream_start_time and self.stream_files_count > 0:
                elapsed = time.time() - self.stream_start_time
                # Dodano rzutowanie na float() oraz 0.0 zamiast 0
                speed = float(self.stream_files_count / elapsed) if elapsed > 0 else 0.0
                self.stream_speed_label.configure(
                    text=f"Prędkość: {speed:.1f} plików/s"
                )
                if len(self.stream_queue) > 0 and speed > 0:
                    eta_seconds = len(self.stream_queue) / speed
                    if eta_seconds < 60:
                        eta_str = f"~{int(eta_seconds)}s"
                    elif eta_seconds < 3600:
                        eta_str = f"~{int(eta_seconds / 60)}min"
                    else:
                        eta_str = f"~{int(eta_seconds / 3600)}h"
                    self.stream_eta_label.configure(text=f"ETA: {eta_str}")
                else:
                    self.stream_eta_label.configure(text="ETA: -")

        self.after(0, _update)

    def clear_stream(self):
        self.stream_queue = []
        self.stream_current = None
        self.stream_completed = []
        self.stream_start_time = None
        self.stream_files_count = 0
        if self.stream_listbox:
            self.stream_listbox.delete(0, tk.END)

    # =================================

    def setup_tab(
            self,
            parent,
            mode,
            src_label_text,
            dst_label_text,
            show_order_button,
            extra_ui_setup=None,
            dashboard=False,
    ):
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
        card.grid(
            row=0, column=0, padx=10, pady=(10, 8), sticky="new"
        )  # mniejsze padx/pady
        card.grid_columnconfigure(1, weight=1)

        # ŹRÓDŁO + HISTORIA
        ctk.CTkLabel(
            card, text=src_label_text, font=font_label, text_color="#E0E0E0"
        ).grid(row=0, column=0, padx=15, pady=(20, 10), sticky="w")
        entry_src = ctk.CTkEntry(
            card, placeholder_text="Wskaż folder...", height=32, border_width=1
        )  # było 36
        entry_src.grid(row=0, column=1, padx=5, pady=(20, 10), sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda e=entry_src: self.select_dir(e),
            width=100,
            height=32,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=0, column=2, padx=(15, 5), pady=(20, 10))
        hist_btn_src = ctk.CTkButton(
            card,
            text="🕒",
            width=36,
            height=36,
            fg_color="#333333",
            hover_color="#444444",
            font=ctk.CTkFont(size=16),
        )
        hist_btn_src.grid(row=0, column=3, padx=(0, 15), pady=(20, 10))
        hist_btn_src.bind(
            "<Button-1>", lambda event, e=entry_src: self.show_history_menu(event, e)
        )
        add_tooltip(hist_btn_src, "Pokaż historię ostatnio używanych folderów")

        # CEL + HISTORIA
        pady_bottom = (0, 10) if extra_ui_setup else (0, 20)
        ctk.CTkLabel(
            card, text=dst_label_text, font=font_label, text_color="#E0E0E0"
        ).grid(row=1, column=0, padx=15, pady=pady_bottom, sticky="w")
        entry_dst = ctk.CTkEntry(
            card, placeholder_text="Wskaż lokalizację...", height=36, border_width=1
        )
        entry_dst.grid(row=1, column=1, padx=5, pady=pady_bottom, sticky="ew")
        ctk.CTkButton(
            card,
            text="Przeglądaj",
            image=self.icon_folder,
            command=lambda e=entry_dst: self.select_dir(e),
            width=110,
            height=36,
            font=font_btn,
            fg_color="#333333",
            hover_color="#444444",
        ).grid(row=1, column=2, padx=(15, 5), pady=pady_bottom)
        hist_btn_dst = ctk.CTkButton(
            card,
            text="🕒",
            width=36,
            height=36,
            fg_color="#333333",
            hover_color="#444444",
            font=ctk.CTkFont(size=16),
        )
        hist_btn_dst.grid(row=1, column=3, padx=(0, 15), pady=pady_bottom)
        hist_btn_dst.bind(
            "<Button-1>", lambda event, e=entry_dst: self.show_history_menu(event, e)
        )
        add_tooltip(hist_btn_dst, "Pokaż historię ostatnio używanych folderów")

        if extra_ui_setup:
            extra_ui_setup(card, 2)

        # DASHBOARD JEŚLI WŁĄCZONY
        current_row = 1
        if dashboard:
            self.dashboard_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            self.dashboard_frame.grid(
                row=current_row, column=0, padx=20, pady=(0, 10), sticky="ew"
            )
            self.build_dashboard_ui(self.dashboard_frame)
            current_row += 1

        btn_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        btn_frame.grid(row=current_row, column=0, padx=20, pady=(5, 20), sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)
        btn = ctk.CTkButton(
            btn_frame,
            text="Rozpocznij proces",
            image=self.icon_start,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#0067C0",
            hover_color="#005A9E",
            height=40,
            corner_radius=6,
            command=lambda m=mode: self.start_pipeline(m),
        )
        if show_order_button:
            btn.pack(side="left", expand=True, fill="x", padx=(0, 10))
            cfg_btn = ctk.CTkButton(
                btn_frame,
                text="Skonfiguruj układ PDF",
                font=font_btn,
                height=44,
                corner_radius=6,
                fg_color="transparent",
                border_width=1,
                border_color="#555555",
                hover_color="#333333",
                command=lambda m=mode, e=entry_dst: self.open_mode_order_window(
                    m, e.get().strip()
                ),
            )
            cfg_btn.pack(side="right")
            add_tooltip(
                cfg_btn,
                "Pozwala ustalić w jakiej kolejności ułożą się dokumenty wejściowe w finalnym dokumencie PDF.",
            )
        else:
            btn.pack(side="left", expand=True, fill="x")
        self.entries[mode] = {"src": entry_src, "dst": entry_dst, "btn": btn}

    def select_dir(self, entry_widget):
        folder = filedialog.askdirectory()
        if folder:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, folder)
            self.add_to_history(folder)
            entry_name = None
            for attr in dir(self):
                if getattr(self, attr, None) is entry_widget:
                    entry_name = attr
                    break
            if entry_name:
                self.set_setting(f"folder_{entry_name}", folder)

    def select_file(self, entry_widget, filetypes):
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, file_path)

    def select_save_file(self, entry_widget):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Dokument Word", "*.docx")],
            title="Zapisz szablon jako",
        )
        if file_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, file_path)

    def log(self, text):
        def _update_log():
            self.textbox.configure(state="normal")
            self.textbox.insert("end", text + "\n")
            self.textbox.see("end")
            self.textbox.configure(state="disabled")

        self.after(0, _update_log)

    def show_validation_window_sync(self, title_text, warnings):
        """ Pokazuje okno walidacji i blokuje wątek w tle do momentu decyzji użytkownika. """
        proceed_evt = threading.Event()
        cancel_evt = threading.Event()

        self.after(0, lambda: ValidationWindow(self, title_text, warnings, proceed_evt, cancel_evt))

        # Blokada wątku dopóki nie zostanie wciśnięty żaden z przycisków
        while not proceed_evt.is_set() and not cancel_evt.is_set():
            time.sleep(0.1)

        return proceed_evt.is_set()

    def set_progress(self, value, current_file=None, current=None, total=None, description=None):
        if total is not None:
            self.progress_total = max(0, int(total))

        if current is not None:
            self.progress_current = max(0, int(current))
        elif value is not None and self.progress_total:
            self.progress_current = int(round(float(value) * self.progress_total))

        if current_file is not None:
            self.progress_current_file = str(current_file)

        if description is not None:
            self.progress_description = str(description)

        try:
            bar_value = max(0.0, min(1.0, float(value)))
        except Exception:
            bar_value = 0.0
        self.progress_bar_value = bar_value

        def _update():
            try:
                self.progress_bar.set(bar_value)
                self.progress_percent_label.configure(text=f"{int(round(bar_value * 100))}%")
                self.progress_detail_label.configure(text=self._build_progress_detail())
                self.progress_eta_label.configure(text=self._calculate_progress_eta())
            except Exception:
                pass

        self.after(0, _update)

    def start_progress_tracking(self, total, description=""):
        self.progress_total = max(0, int(total))
        self.progress_current = 0
        self.progress_start_time = time.time()
        self.progress_current_file = None
        self.progress_description = description
        self.set_progress(0, current=0, total=self.progress_total, description=description)

    def reset_progress_details(self, text="Oczekiwanie na zadanie"):
        self.progress_total = 0
        self.progress_current = 0
        self.progress_start_time = None
        self.progress_current_file = None
        self.progress_description = ""

        def _reset():
            try:
                self.progress_bar.set(0)
                self.progress_percent_label.configure(text="0%")
                self.progress_detail_label.configure(text=text)
                self.progress_eta_label.configure(text="")
            except Exception:
                pass

        self.after(0, _reset)

    def _build_progress_detail(self):
        parts = []

        if getattr(self, "progress_description", ""):
            parts.append(self.progress_description)

        if getattr(self, "progress_total", 0) > 0:
            parts.append(
                f"Przetwarzanie: {getattr(self, 'progress_current', 0)} / {self.progress_total}"
            )

        if getattr(self, "progress_current_file", None):
            parts.append(f"Plik: {self.progress_current_file}")

        return "   |   ".join(parts) if parts else "Oczekiwanie na zadanie"

    def _calculate_progress_eta(self):
        total = getattr(self, "progress_total", 0)
        current = getattr(self, "progress_current", 0)
        start_time = getattr(self, "progress_start_time", None)

        if not total or current <= 0 or not start_time:
            return ""

        elapsed = time.time() - start_time
        if elapsed <= 0:
            return ""

        rate = current / elapsed
        if rate <= 0:
            return ""

        remaining = max(0, total - current)
        eta_seconds = remaining / rate

        return f"Pozostało: {self._format_duration(eta_seconds)}"

    @staticmethod
    def _format_duration(seconds):
        seconds = max(0, int(round(seconds)))

        if seconds < 60:
            return f"~{seconds} s"

        if seconds < 3600:
            return f"~{seconds // 60} min"

        hours = seconds // 3600
        minutes = (seconds % 3600) // 60

        if minutes:
            return f"~{hours} h {minutes} min"

        return f"~{hours} h"

    def open_mode_order_window(self, mode_key, output_root):
        if not output_root:
            messagebox.showwarning(
                "Wymagana konfiguracja",
                "Wskaż najpierw lokalizację docelową, aby zapisać układ dla tego profilu.",
            )
            return
        out_root = Path(output_root)
        out_root.mkdir(parents=True, exist_ok=True)
        config_folder = out_root / "PDF"
        config_folder.mkdir(parents=True, exist_ok=True)
        PdfOrderWindow(self, config_folder, mode_key)

    def open_manual_merge_window(self):
        src = self.manual_pdf_src.get().strip()
        dst = self.manual_pdf_dst.get().strip()
        if not src or not Path(src).exists():
            messagebox.showwarning(
                "Brak danych",
                "Sprawdź, czy podano prawidłowy folder zawierający pliki PDF.",
            )
            return
        if not dst:
            messagebox.showwarning("Brak danych", "Wskaż prawidłowy folder docelowy.")
            return
        ManualPdfMergeWindow(self, Path(src), Path(dst))

    def _disable_ui_for_process(self):
        self.running = True
        self.stop_event.clear()
        self.start_progress_tracking(0, "Przygotowywanie zadania...")
        self.stop_btn.configure(state="normal", text="Przerwij zadanie")
        self.open_dir_btn.configure(state="disabled")
        for m in self.entries:
            self.entries[m]["btn"].configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.excel_start_btn is not None:
            self.excel_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        # --- ZABLOKOWANIE PRZYCISKU USUWANIA ---
        if hasattr(self, 'remove_cols_btn') and self.remove_cols_btn is not None:
            self.remove_cols_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if hasattr(self, 'cb_remove_owners'):
            self.cb_remove_owners.configure(state="disabled")
            self.cb_remove_ls.configure(state="disabled")

        if self.title_generate_btn is not None:
            self.title_generate_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.mietek_title_generate_btn is not None:
            self.mietek_title_generate_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.layout_merge_btn is not None:
            self.layout_merge_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.pdfconv_start_btn is not None:
            self.pdfconv_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.split_pdf_btn is not None:
            self.split_pdf_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.mdb_start_btn is not None:
            self.mdb_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.rozl_start_btn is not None:
            self.rozl_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.mietki_start_btn is not None:
            self.mietki_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.krzyz_start_btn is not None:
            self.krzyz_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.halizny_start_btn is not None:
            self.halizny_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        if self.excel_z_mdb_start_btn is not None:
            self.excel_z_mdb_start_btn.configure(
                state="disabled", text="Przetwarzanie...", fg_color="#444444"
            )
        for mode in self.tpl_data:
            if "btn_gen" in self.tpl_data[mode]:
                self.tpl_data[mode]["btn_gen"].configure(
                    state="disabled", text="Przetwarzanie...", fg_color="#444444"
                )
        if self.stream_frame:
            self.stream_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        if hasattr(self, 'mietek_rozb_start_btn') and self.mietek_rozb_start_btn is not None:
            self.mietek_rozb_start_btn.configure(state="disabled", text="Przetwarzanie...", fg_color="#444444")
        if hasattr(self, 'mietek_rozb_bez_nazwisk_btn') and self.mietek_rozb_bez_nazwisk_btn is not None:
            self.mietek_rozb_bez_nazwisk_btn.configure(state="disabled", text="Przetwarzanie...", fg_color="#444444")
        if hasattr(self, 'nazwiska_mietek_start_btn') and self.nazwiska_mietek_start_btn is not None:
            self.nazwiska_mietek_start_btn.configure(state="disabled", text="Przetwarzanie...", fg_color="#444444")

    def check_stop(self):
        if self.stop_event.is_set():
            raise InterruptedError()

    # NOWA METODA: Generowanie STR_TYT w trybie Pełny Automat
    def restore_all_buttons(self):
        self.stop_btn.configure(state="disabled", text="Przerwij zadanie")
        if self.last_output_dir and Path(self.last_output_dir).exists():
            self.open_dir_btn.configure(state="normal")
        else:
            self.open_dir_btn.configure(state="disabled")
        for m in self.entries:
            self.entries[m]["btn"].configure(
                state="normal", text="Rozpocznij proces", fg_color="#0067C0"
            )
        if self.excel_start_btn is not None:
            self.excel_start_btn.configure(
                state="normal", text="Uruchom układanie Exceli", fg_color="#0067C0"
            )
        # --- ODBLOKOWANIE PRZYCISKU I CHECKBOXÓW USUWANIA ---
        if hasattr(self, 'remove_cols_btn') and self.remove_cols_btn is not None:
            self.remove_cols_btn.configure(
                state="normal", text="Usuń kolumny (wg zaznaczenia)", fg_color="#8B0000"
            )
        if hasattr(self, 'cb_remove_owners'):
            self.cb_remove_owners.configure(state="normal")
            self.cb_remove_ls.configure(state="normal")

        if self.title_generate_btn is not None:
            self.title_generate_btn.configure(
                state="normal", text="Masowo twórz strony STR_TYT", fg_color="#0067C0"
            )
        if self.mietek_title_generate_btn is not None:
            self.mietek_title_generate_btn.configure(
                state="normal", text="Masowo twórz strony STR_TYT", fg_color="#0067C0"
            )
        if self.layout_merge_btn is not None:
            self.layout_merge_btn.configure(
                state="normal", text="Twórz gotowe PDF", fg_color="#0067C0"
            )
        if self.pdfconv_start_btn is not None:
            self.pdfconv_start_btn.configure(
                state="normal", text="Konwertuj wszystko do PDF", fg_color="#0067C0"
            )
        if self.split_pdf_btn is not None:
            self.split_pdf_btn.configure(
                state="normal", text="Rozdziel na osobne PDF", fg_color="#0067C0"
            )
        if self.mdb_start_btn is not None:
            self.mdb_start_btn.configure(
                state="normal", text="Usuń 0 w bazach (MDB)", fg_color="#0067C0"
            )
        if self.rozl_start_btn is not None:
            self.rozl_start_btn.configure(
                state="normal", text="Uruchom rozliczanie obrębów", fg_color="#0067C0"
            )
        if self.mietki_start_btn is not None:
            self.mietki_start_btn.configure(
                state="normal", text="Klonuj strukturę folderów", fg_color="#0067C0"
            )
        if self.krzyz_start_btn is not None:
            self.krzyz_start_btn.configure(
                state="normal", text="Wstrzyknij krzyżówki do DBF", fg_color="#0067C0"
            )
        if self.halizny_start_btn is not None:
            self.halizny_start_btn.configure(
                state="normal", text="Przenieś halizny w D*.DBF", fg_color="#0067C0"
            )
        if self.excel_z_mdb_start_btn is not None:
            self.excel_z_mdb_start_btn.configure(
                state="normal", text="Wyciągnij dane z MDB", fg_color="#0067C0"
            )
        for mode in self.tpl_data:
            if "btn_gen" in self.tpl_data[mode]:
                self.tpl_data[mode]["btn_gen"].configure(
                    state="normal", text="Wygeneruj Szablon STR_TYT", fg_color="#27ae60"
                )
        if hasattr(self, 'mietek_rozb_start_btn') and self.mietek_rozb_start_btn is not None:
            self.mietek_rozb_start_btn.configure(state="normal", text="Generuj Wykaz Rozbieżności", fg_color="#0067C0")
        if hasattr(self, 'mietek_rozb_bez_nazwisk_btn') and self.mietek_rozb_bez_nazwisk_btn is not None:
            self.mietek_rozb_bez_nazwisk_btn.configure(state="normal", text="Bez Nazwisk", fg_color="#8B0000")
        if hasattr(self, 'nazwiska_mietek_start_btn') and self.nazwiska_mietek_start_btn is not None:
            self.nazwiska_mietek_start_btn.configure(state="normal", text="Generuj struktury (tylko Ewidencja)",
                                                     fg_color="#0067C0")

        def _finish_progress():
            try:
                self.progress_current_file = None
                self.progress_eta_label.configure(text="")

                if getattr(self, "progress_total", 0) and getattr(self, "progress_current", 0) >= self.progress_total:
                    self.progress_detail_label.configure(text="Zakończono")
            except Exception:
                pass

        self.after(0, _finish_progress)

        if self.stream_frame:
            self.after(3000, lambda: self.stream_frame.grid_remove())
            self.clear_stream()

    # ==========================================
    # ZAKŁADKA: TWORZENIE MIETKÓW
    # ==========================================
