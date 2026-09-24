# -*- coding: utf-8 -*-
"""
Forestly — backend interfejsu webowego (PyWebView).

WebBackend dziedziczy po WSZYSTKICH mixinach zakładek (tak samo jak
ModernApp), ale zamiast okna CustomTkinter udostępnia mostek JS:
  * kontrolki ze schematu (web_schema) tworzą "sztuczne widgety"
    (FakeEntry/FakeVar), które czyta istniejąca logika mixinów,
  * metody usługowe GUI (log, update_status, set_progress, dashboard...)
    zamieniają się w zdarzenia kolejkowane i odpytywane przez frontend
    (api.poll()),
  * okienka messagebox (showwarning/askyesno...) działają przez mostek
    do dialogów JS.

Dzięki temu CAŁA funkcjonalność programu pozostaje bez zmian — zmienia
się tylko warstwa prezentacji.
"""

import json
import os
import sys
import threading
import time
import traceback
import tempfile
import subprocess
from collections import deque
from pathlib import Path

from app.config import (CURRENT_VERSION, SETTINGS_FILE, HISTORY_FILE,
                        load_margins, save_margins,
                        get_saved_template_order, set_saved_template_order,
                        get_saved_excluded_templates, set_saved_excluded_templates,
                        get_default_template_keys, PDF_ORDER_TEMPLATES)

from app.gui.web_schema import build_schema

# Mixiny — identyczny skład jak ModernApp (bez ctk.CTk)
from app.gui.tabs.tab_all import TabAllMixin
from app.gui.tabs.tab_opis_og import TabOpisOgMixin
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
from app.updater import UpdaterMixin

# Filtry plików dla przeglądarek (id kontrolki → file_types dla pywebview)
BROWSE_FILTERS = {
    "all_template": ("Dokument Word", ("*.docx",)),
    "all_skroty": ("Word i PDF", ("*.docx", "*.doc", "*.pdf")),
    "mt_template": ("Dokument Word", ("*.docx", "*.doc")),
    "tt_template": ("Dokument Word", ("*.docx", "*.doc")),
    "zm_src": ("Baza Access", ("*.mdb",)),
}
MARGIN_FILE_TYPES = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "HALIZNY",
                     "WYK_NEG", "OPIS", "ZEST1", "WK_ZM1"]


# --------------------------------------------------------------- sztuczne widgety

class FakeEntry:
    """Zamiennik CTkEntry — przechowuje tekst."""

    def __init__(self, value=""):
        self.value = str(value)

    def get(self):
        return self.value

    def set(self, value):
        self.value = "" if value is None else str(value)

    def insert(self, _pos, value):
        self.set(value)

    def delete(self, *_a):
        self.value = ""

    def configure(self, **_kw):
        pass


class FakeVar:
    """Zamiennik BooleanVar/StringVar."""

    def __init__(self, value=False):
        self._v = value

    def get(self):
        return self._v

    def set(self, value):
        self._v = value


class FakeButton:
    """Zamiennik CTkButton — wszystkie akcje to no-op."""

    def configure(self, **_kw):
        pass


# --------------------------------------------------------------------- backend

class WebBackend(
    TabOpisOgMixin,
    TabAllMixin, TabWordMixin, TabPdfMixin, TabManualMergeMixin,
    TabTemplateGeneratorMixin, TabTitlePagesMixin, TabExcelMixin,
    TabLayoutExcelMixin, TabSplitPdfMixin, TabMdbUpdateMixin,
    TabPdfConverterMixin, TabRozliczanieMixin, TabHaliznyMixin,
    TabWydrukiMixin, TabExcelZMdbMixin, TabTworzenieMietkowMixin,
    TabNazwiskaMietekMixin, TabMietekRozbieznosciMixin, UpdaterMixin,
):
    """Logika aplikacji bez CustomTkinter — z mostkiem do PyWebView."""

    def __init__(self):
        # --- inicjalizacja atrybutów jak w ModernApp.__init__ ---
        self.stop_event = threading.Event()
        self.running = False
        self.last_output_dir = None
        self.entries = {}
        self.tpl_data = {"MIETEK": {}, "TAKSATOR": {}}
        self.margin_vars = {}
        self.wydruki_filter_vars = {}
        self.wydruki_filter_checkboxes = {}
        self.word_filter_vars = {}
        self.word_filter_checkboxes = {}
        self.excel_font_entries = {}

        # zdarzenia do frontendu + dialogi
        self._events = deque()
        self._ev_lock = threading.Lock()
        self._dialog_waits = {}   # id -> (Event, result)

        # widgety (przypisujemy None jak ModernApp, potem fakes ze schematu)
        for attr in ("all_skroty_entry", "all_gdos_entry", "all_tpl_doc_var",
                     "all_tpl_prefix_var", "all_tpl_woj_var", "all_tpl_powiat_var",
                     "all_tpl_gmina_var", "all_tpl_stan_na_entry", "all_tpl_okres_entry",
                     "all_wsk_od_entry", "all_wsk_do_entry",
                     "excel_folder_entry", "excel_output_entry",
                     "excel_start_btn", "title_template_entry", "title_excel_entry",
                     "title_output_entry", "title_village_placeholder_entry",
                     "title_area_placeholder_entry", "title_generate_btn",
                     "mietek_title_template_entry", "mietek_title_word_entry",
                     "mietek_title_village_placeholder_entry",
                     "mietek_title_area_placeholder_entry",
                     "mietek_title_generate_btn",
                     "layout_title_folder_entry", "layout_opisy_folder_entry",
                     "layout_raporty_folder_entry", "layout_output_folder_entry",
                     "layout_merge_btn", "pdfconv_source_entry", "pdfconv_output_entry",
                     "pdfconv_start_btn", "split_title_folder_entry",
                     "split_opisy_folder_entry", "split_raporty_folder_entry",
                     "split_output_folder_entry", "split_pdf_btn",
                     "mdb_source_entry", "mdb_output_entry", "mdb_start_btn",
                     "rozl_xls_entry", "rozl_val_entry", "rozl_out_entry",
                     "rozl_start_btn", "mietki_base_entry", "mietki_out_entry",
                     "mietki_names_textbox", "mietki_start_btn",
                     "nazwiska_bazowy_entry", "nazwiska_out_entry",
                     "nazwiska_mietek_start_btn", "krzyz_xls_entry",
                     "krzyz_mietki_entry", "krzyz_start_btn",
                     "halizny_mietki_entry", "halizny_start_btn",
                     "excel_z_mdb_src_entry", "excel_z_mdb_out_entry",
                     "excel_z_mdb_start_btn", "mietek_rozb_mietki_entry",
                     "mietek_rozb_excel_entry", "mietek_rozb_out_entry",
                     "mietek_rozb_start_btn", "mietek_rozb_bez_nazwisk_btn",
                     "wydruki_mietki_entry", "wydruki_all_btn",
                     "manual_pdf_src", "manual_pdf_dst",
                     "opis_og_root_entry",
                     "opis_og_taksator_entry", "opis_og_taksator_gdos_entry",
                     "opis_og_taksator_start_btn", "gdos_xlsx_entry"):
            setattr(self, attr, None)

        self._install_messagebox_shim()
        self._build_fakes_from_schema()
        self._check_pending_changelog()
        # v2.0.6: automatyczne sprawdzenie nowej wersji po starcie
        # (odpowiednik self.after(2000, ...) z klasycznego GUI)
        self.after(2000, lambda: self.check_github_update(manual=False))

    def _check_pending_changelog(self):
        """Po aktualizacji: pokazuje changelog zapisany przez aktualizator
        (pending_changelog.json obok EXE) — odpowiednik check_pending_changelog
        z klasycznego GUI."""
        try:
            if getattr(sys, "frozen", False):
                app_dir = Path(sys.executable).resolve().parent
            else:
                app_dir = Path(__file__).resolve().parent
            changelog_file = app_dir / "pending_changelog.json"
            if not changelog_file.exists():
                return
            data = json.loads(changelog_file.read_text(encoding="utf-8"))
            version = data.get("version", CURRENT_VERSION)
            body = data.get("changelog", "")
            try:
                changelog_file.unlink()
            except Exception:
                pass
            if body.strip():
                self._emit({"type": "changelog", "version": version, "body": body})
                self.log(f"Zainstalowano aktualizację {version} — zobacz, co nowego.")
        except Exception as e:
            print(f"[INFO] Błąd odczytu changelogu: {e}")

    # -------------------------------------------------- zdarzenia i dialogi

    def _emit(self, event):
        with self._ev_lock:
            self._events.append(event)

    def _install_messagebox_shim(self):
        """Zamienia tkinter.messagebox na dialogi JS (wywoływane z wątków)."""
        import tkinter.messagebox as _mb

        def _fire(kind, title, message):
            self._emit({"type": "dialog", "kind": kind, "title": str(title),
                        "message": str(message)})

        def _ask(title, message):
            did = f"d{time.time():.6f}".replace(".", "")
            ev = threading.Event()
            self._dialog_waits[did] = (ev, [False])
            self._emit({"type": "dialog", "kind": "confirm", "id": did,
                        "title": str(title), "message": str(message)})
            ev.wait(timeout=3600)
            return bool(self._dialog_waits.pop(did, (None, [False]))[1][0])

        _mb.showinfo = lambda t, m, **k: _fire("info", t, m)
        _mb.showwarning = lambda t, m, **k: _fire("warn", t, m)
        _mb.showerror = lambda t, m, **k: _fire("error", t, m)
        _mb.askyesno = lambda t, m, **k: _ask(t, m)
        _mb.askokcancel = lambda t, m, **k: _ask(t, m)

    def dialog_reply(self, dialog_id, answer):
        wait = self._dialog_waits.get(dialog_id)
        if wait:
            wait[1][0] = bool(answer)
            wait[0].set()
        return {"ok": True}

    # ------------------------------------------------- sztuczne widgety

    def _navigate(self, path, create_dicts=True):
        """Rozwiązuje ścieżkę atrybutu 'entries.ALL.src' do miejsca zapisu."""
        parts = path.split(".")
        obj = self
        for p in parts[:-1]:
            if isinstance(obj, dict):
                obj = obj.setdefault(p, {}) if create_dicts else obj.get(p)
            else:
                nxt = getattr(obj, p, None)
                if nxt is None and create_dicts:
                    nxt = {}
                    setattr(obj, p, nxt)
                obj = nxt
        return obj, parts[-1]

    def _set_fake(self, attr, fake):
        obj, key = self._navigate(attr)
        if isinstance(obj, dict):
            obj[key] = fake
        else:
            setattr(obj, key, fake)

    def _get_fake(self, attr):
        if "." not in attr:
            return getattr(self, attr, None)
        obj, key = attr.rsplit(".", 1)
        holder = self
        for p in obj.split("."):
            if isinstance(holder, dict):
                holder = holder.get(p)
            else:
                holder = getattr(holder, p, None)
            if holder is None:
                return None
        if isinstance(holder, dict):
            return holder.get(key)
        return getattr(holder, key, None)

    def _build_fakes_from_schema(self):
        self.schema = build_schema()
        for tab in self.schema["tabs"]:
            self._fakes_for_controls(tab["controls"])
        # kompatybilność: pola WSIE.DBF czytane też z klasycznych kluczy
        for pref, key in (("wsie_wojew", "wsie_wojew"), ("wsie_powiat", "wsie_powiat"),
                          ("wsie_stan", "wsie_stan"), ("wsie_obod", "wsie_obod"),
                          ("wsie_obdo", "wsie_obdo"), ("wsie_nrws", "wsie_nrws"),
                          ("wsie_rokz", "wsie_rokz"),
                          ("nm_wsie_wojew", "wsie_wojew"), ("nm_wsie_powiat", "wsie_powiat"),
                          ("nm_wsie_stan", "wsie_stan"), ("nm_wsie_obod", "wsie_obod"),
                          ("nm_wsie_obdo", "wsie_obdo"), ("nm_wsie_nrws", "wsie_nrws"),
                          ("nm_wsie_rokz", "wsie_rokz")):
            fake = self._get_fake(f"{pref}_entry") if "." not in pref else None
            if fake is not None:
                fake.set(self.get_setting(key, fake.get()))

        # przyciski — uniwersalne atrapy (sterowanie widocznością robi JS)
        for tab in self.schema["tabs"]:
            for b in tab["buttons"]:
                pass  # przyciski istnieją tylko w frontendcie

        # remove_names_var potrzebny dla ALL/WORD
        if not hasattr(self, "remove_names_var"):
            self.remove_names_var = FakeVar(
                bool(self.get_setting("web.remove_names", True)))


    def _fakes_for_controls(self, controls):
        """Tworzy atrapy widgetów dla kontrolek schematu (wraz z grupami)."""
        for c in controls:
            kind = c["kind"]
            if kind == "group":
                self._fakes_for_controls(c.get("controls") or [])
                continue
            if kind in ("path", "text"):
                val = self.get_setting(f"web.{c['id']}", c.get("default", "") or "")
                self._set_fake(c["attr"], FakeEntry(val))
            elif kind == "check":
                val = self.get_setting(f"web.{c['id']}", c.get("default", False))
                self._set_fake(c["attr"], FakeVar(bool(val)))
            elif kind == "checks":
                base = c["attr_base"]
                for choice in c["choices"]:
                    var = FakeVar(choice == "Wszystkie")
                    self._set_fake(f"{base}.{choice}", var)
            elif kind == "select":
                self._set_fake(c["attr"], FakeVar(c.get("default", "")))
            elif kind == "margins":
                mode = c["mode"]
                saved = load_margins().get(mode, {})
                for ftype in MARGIN_FILE_TYPES:
                    fsaved = saved.get(ftype, {})
                    for side, dflt in (("T", "1.5"), ("B", "1.5"),
                                       ("L", "2.5"), ("R", "1.5")):
                        self._set_fake(
                            f"margin_vars.{mode}.{ftype}.{side}",
                            FakeEntry(str(fsaved.get(side, dflt))))
            elif kind == "fonts":
                for f in c["fonts"]:
                    self.excel_font_entries[f["sheet"]] = {
                        "entry": FakeEntry(str(self.get_setting(
                            f"web.font.{f['sheet']}", f["default"]))),
                        "start_row": f["start_row"],
                    }
                self.global_font_entry = FakeEntry(
                    self.get_setting("web.xl_global_size", "10"))



    # ------------------------------------------------ wartości z frontendu

    def set_values(self, values):
        """Przyjmuje słownik {id_kontrolki: wartość} z frontendu."""
        controls = {}
        for tab in self.schema["tabs"]:
            for c in tab["controls"]:
                if c.get("id"):
                    controls[c["id"]] = c
        for cid, val in (values or {}).items():
            c = controls.get(cid)
            if not c:
                continue
            kind = c["kind"]
            if kind in ("path", "text"):
                self._get_fake(c["attr"]).set(val)
            elif kind == "check":
                self._get_fake(c["attr"]).set(bool(val))
            elif kind == "select":
                self._get_fake(c["attr"]).set(str(val))
            elif kind == "checks":
                val = dict(val or {})
                # logika wzajemnego wykluczania (jak w GUI):
                # 'Wszystkie' zaznaczone -> pozostałe odznaczone;
                # cokolwiek innego zaznaczone -> 'Wszystkie' odznaczone.
                if val.get("Wszystkie"):
                    val = {k: (k == "Wszystkie") for k in val}
                elif any(v and k != "Wszystkie" for k, v in val.items()):
                    val["Wszystkie"] = False
                for choice in c["choices"]:
                    f = self._get_fake(f"{c['attr_base']}.{choice}")
                    if f is not None:
                        f.set(bool(val.get(choice, choice == "Wszystkie" and not any(
                            v and k != "Wszystkie" for k, v in val.items()))))
            elif kind == "margins":
                mode_cfg = {}
                all_valid = True
                for ftype, sides in (val or {}).items():
                    mode_cfg[ftype] = {}
                    for side, v in (sides or {}).items():
                        f = self._get_fake(f"margin_vars.{c['mode']}.{ftype}.{side}")
                        if f is not None:
                            f.set(str(v))
                        try:
                            mode_cfg[ftype][side] = float(str(v).replace(",", "."))
                        except (TypeError, ValueError):
                            all_valid = False
                # trwałe zapamiętanie marginesów — jak w GUI CustomTkinter
                # (tam save_margins leci przy starcie zadania; w web zapisujemy
                # od razu, żeby przetrwały restart programu). Niepoprawne
                # wartości (w trakcie wpisywania) nie trafiają na dysk.
                if mode_cfg and all_valid:
                    saved_config = load_margins()
                    saved_config[c["mode"]] = mode_cfg
                    save_margins(saved_config)
            elif kind == "fonts":
                for sheet, v in (val or {}).items():
                    ent = self.excel_font_entries.get(sheet)
                    if ent:
                        ent["entry"].set(str(v))
            if c.get("save"):
                self.set_setting(f"web.{cid}", val)
        return {"ok": True}

    # ------------------------------------------------------- usługi GUI

    def log(self, text):
        self._emit({"type": "log", "text": str(text)})

    def clear_log(self):
        self._emit({"type": "clear_log"})

    def update_status(self, text, color="#0078D7", animate=True):
        self._emit({"type": "status", "text": str(text), "color": color})

    def set_progress(self, value, current_file=None, current=None,
                     total=None, description=None):
        self._emit({"type": "progress", "value": value,
                    "file": current_file, "current": current,
                    "total": total, "desc": description})

    def start_progress_tracking(self, total, description=""):
        self._emit({"type": "progress", "value": 0, "total": total,
                    "desc": description})

    def update_dashboard(self, step_index, status, text=None):
        self._emit({"type": "dashboard", "step": step_index,
                    "status": status, "text": text})

    def reset_dashboard(self):
        self._emit({"type": "dashboard_reset"})

    def build_dashboard_ui(self, parent):
        pass

    def update_options_visibility(self):
        pass

    def after(self, ms, func=None, *args):
        if func is None:
            return None

        def _run():
            try:
                func(*args)
            except Exception:
                pass

        if ms and ms > 0:
            threading.Timer(ms / 1000.0, _run).start()
        else:
            threading.Thread(target=_run, daemon=True).start()

    def check_stop(self):
        if self.stop_event.is_set():
            raise InterruptedError()

    def _disable_ui_for_process(self):
        self.running = True
        self.stop_event.clear()
        self._emit({"type": "state", "running": True})

    def restore_all_buttons(self):
        self.running = False
        self._emit({"type": "state", "running": False})

    def disable_all_buttons(self):
        self._disable_ui_for_process()

    # ------------------------------------------- ustawienia i historia

    def load_settings(self):
        if SETTINGS_FILE.exists():
            try:
                return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def save_settings(self, settings):
        try:
            SETTINGS_FILE.write_text(
                json.dumps(settings, ensure_ascii=False, indent=2),
                encoding="utf-8")
        except Exception:
            pass

    def get_setting(self, key, default=""):
        return self.load_settings().get(key, default)

    def set_setting(self, key, value):
        settings = self.load_settings()
        settings[key] = value
        self.save_settings(settings)

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
                json.dumps(history, ensure_ascii=False, indent=2),
                encoding="utf-8")
        except Exception:
            pass

    def add_to_history(self, path):
        if not path or not Path(path).exists():
            return
        history = [p for p in self.load_history() if p != path]
        history.insert(0, path)
        self.save_history(history[:10])

    def show_history_menu(self, *a, **kw):
        pass

    # -------------------------------------------------- wybór plików/folderów

    def browse(self, control_id, kind):
        """Otwiera natywne okno wyboru (pywebview) i zwraca ścieżkę."""
        try:
            import webview
            win = webview.windows[0]
            # pywebview >= 6: nowe stałe; starsze: fallback na dotychczasowe
            try:
                _fd = webview.FileDialog
                D_FOLDER, D_OPEN, D_SAVE = _fd.FOLDER, _fd.OPEN, _fd.SAVE
            except Exception:
                D_FOLDER = getattr(webview, "FOLDER_DIALOG", None)
                D_OPEN = getattr(webview, "OPEN_DIALOG", None)
                D_SAVE = getattr(webview, "SAVE_DIALOG", None)
            if kind == "folder":
                result = win.create_file_dialog(D_FOLDER)
                path = result[0] if result else None
            elif kind == "save":
                result = win.create_file_dialog(
                    D_SAVE,
                    save_filename="Dokument.xlsx" if control_id == "zm_out"
                    else ("Szablon.docx" if "tpl_" in control_id else "Plik.pdf"))
                path = result if isinstance(result, str) else (result[0] if result else None)
            else:
                fdesc = BROWSE_FILTERS.get(control_id, ("Wszystkie pliki", ("*.*",)))
                result = win.create_file_dialog(
                    D_OPEN,
                    # pywebview wymaga wzorców rozdzielonych ŚREDNIKAMI (*.docx;*.doc)
                file_types=(f"{fdesc[0]} ({';'.join(fdesc[1])})",),)
                path = result[0] if result else None
        except Exception as e:
            self.log(f"[BŁĄD] Wybór pliku: {e}")
            path = None
        if path:
            # folder trafia do historii wprost; plik — jego katalog nadrzędny
            self.add_to_history(str(path) if kind == "folder" else str(Path(path).parent))
            controls = {c["id"]: c for t in self.schema["tabs"]
                        for c in t["controls"] if c.get("id")}
            c = controls.get(control_id)
            if c and c["kind"] == "path":
                self._get_fake(c["attr"]).set(str(path))
            if c and c.get("save"):
                self.set_setting(f"web.{control_id}", str(path))
        return {"path": str(path) if path else None}

    def get_history(self):
        return {"history": self.load_history()}

    # ---------------------------------------------------------- uruchamianie

    def _task_map(self):
        return {
            "start_pipeline:ALL": lambda: self.start_pipeline("ALL"),
            "start_pipeline:WORD": lambda: self.start_pipeline("WORD"),
            "start_pipeline:PDF": lambda: self.start_pipeline("PDF"),
            "start_wydruki_all": self.start_wydruki_all,
            "start_halizny": self.start_halizny_pipeline,
            "generate_template:MIETEK": lambda: self.generate_template_now("MIETEK"),
            "generate_template:TAKSATOR": lambda: self.generate_template_now("TAKSATOR"),
            "start_mietek_title_pages": self.start_mietek_title_pages_pipeline,
            "start_title_pages": self.start_title_pages_pipeline,
            "start_rozbieznosci": lambda: self.start_mietek_rozbieznosci_pipeline(bez_nazwisk=False),
            "start_rozbieznosci_bez": lambda: self.start_mietek_rozbieznosci_pipeline(bez_nazwisk=True),
            "start_nazwiska_mietek": self.start_nazwiska_mietek_pipeline,
            "start_excel": self.start_excel_pipeline,
            "start_layout_excel": self.start_layout_excel_pipeline,
            "start_split_pdf": self.start_split_pdf_pipeline,
            "start_opis_og": self.start_opis_og_pipeline,
            "start_opis_og_taksator": self.start_opis_og_taksator_pipeline,
            "gdos_export": self.gdos_export_task,
            "gdos_import": self.gdos_import_task,
            "start_mdb_update": self.start_mdb_update_pipeline,
            "start_excel_z_mdb": self.start_excel_z_mdb_pipeline,
            "start_pdf_converter": self.start_pdf_converter_pipeline,
            "start_rozliczanie": self.start_rozliczanie_pipeline,
            "start_zestawienie": self.start_zestawienie_zbiorcze,
            "start_zestawienie_mietki": self.start_zestawienie_mietki,
            "start_tworzenie_mietkow": self.start_tworzenie_mietkow_pipeline,
            "start_mietki_krzyzowki": self.start_mietki_i_krzyzowki_pipeline,
        }

    def run(self, task_id):
        task = self._task_map().get(task_id)
        if task is None:
            return {"ok": False, "error": f"Nieznane zadanie: {task_id}"}
        if self.running:
            return {"ok": False, "error": "Zadanie już trwa — zatrzymaj bieżące."}

        def _wrap():
            try:
                task()
            except Exception:
                self.log(traceback.format_exc())
                self.update_status("Błąd", "#D83B01", animate=False)
                self.running = False
                self._emit({"type": "state", "running": False})
            # jeśli start_* odpalił własny wątek — running spadnie przez restore

        threading.Thread(target=_wrap, daemon=True).start()
        return {"ok": True}

    def stop(self):
        self.stop_event.set()
        self.log("[STOP] Zatrzymywanie po bieżącym kroku...")
        return {"ok": True}

    def open_last_output(self):
        """Otwiera Eksplorator w folderze wyników ostatniego zadania."""
        try:
            d = getattr(self, "last_output_dir", None)
            if not (d and Path(d).exists()):
                return {"ok": False,
                        "error": "Folder wyników nieznany \u2014 uruchom najpierw zadanie."}
            if os.name == "nt":
                os.startfile(str(d))  # tylko Windows
            else:
                self.log("[UWAGA] Otwieranie folderu jest dostępne na Windows.")
            return {"ok": True, "path": str(d)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def force_reset(self):
        self.running = False
        self.stop_event.clear()
        self._emit({"type": "state", "running": False})
        return {"ok": True}

    # --------------------------------------- baza obszarów GDOŚ (edytor web)
    def _gdos_plik_zapisu(self):
        """Plik zapisu bazy: obok EXE (wersja przenośna) albo repo root (dev)."""
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent / "gdos_obszary.json"
        return Path(__file__).resolve().parent.parent.parent / "gdos_obszary.json"

    def gdos_list(self):
        """Lista obszarów ochrony przyrody z gdos_obszary.json (dla edytora)."""
        try:
            from app.core.word_worker import get_resource_path
            with open(get_resource_path("gdos_obszary.json"), encoding="utf-8") as f:
                rows = json.load(f).get("obszary", [])
            return {"ok": True, "rows": rows}
        except Exception:
            return {"ok": False, "error": "Nie można wczytać pliku gdos_obszary.json"}

    def gdos_save(self, rows_json):
        """Zapis edytowanej bazy obszarów + odświeżenie cache generatora."""
        try:
            rows = json.loads(rows_json) if isinstance(rows_json, str) else rows_json
            if not isinstance(rows, list):
                raise ValueError("nieprawidłowy format danych")
            czyste = []
            for r in rows:
                nazwa = str(r.get("nazwa", "")).strip()
                if not nazwa:
                    continue
                czyste.append({k: str(r.get(k, "") or "").strip()
                               for k in ("nazwa", "typ", "kod", "pzo",
                                         "powiazanie", "opis")})
            # zachowaj opis pliku z oryginału
            opis = ("Baza obszarów ochrony przyrody dla generatora opisów "
                    "ogólnych (folder z wynikami GDOŚ).")
            try:
                from app.core.word_worker import get_resource_path
                with open(get_resource_path("gdos_obszary.json"), encoding="utf-8") as f:
                    opis = json.load(f).get("_opis", opis)
            except Exception:
                pass
            sciezka = self._gdos_plik_zapisu()
            with open(sciezka, "w", encoding="utf-8") as f:
                json.dump({"_opis": opis, "obszary": czyste},
                          f, ensure_ascii=False, indent=1)
            # cache bazy w generatorze opisów musi się odświeżyć
            from app.gui.tabs.tab_opis_og import TabOpisOgMixin
            TabOpisOgMixin._gdos_kb_cache = None
            self.log("[OK] Baza obszarów GDOŚ zapisana: "
                     + str(len(czyste)) + " obszarów → " + str(sciezka))
            return {"ok": True, "count": len(czyste)}
        except Exception:
            self.log("[BŁĄD] Zapis bazy obszarów GDOŚ:\n"
                     + traceback.format_exc())
            return {"ok": False, "error": "Nie udało się zapisać bazy (szczegóły w logu)"}

    # --------------------------------------- baza GDOŚ: Excel import/eksport (web)
    def gdos_export_task(self):
        """Eksport bazy GDOŚ do Excela (plik obok programu)."""
        try:
            out = self._gdos_plik_zapisu().parent / "gdos_obszary.xlsx"
            n = self.gdos_exportuj_excel(out)
            self.log(f"[OK] Wyeksportowano bazę GDOŚ ({n} obszarów) → {out}")
            self.update_status("Baza wyeksportowana", "#107C10", animate=False)
            return {"ok": True, "path": str(out)}
        except Exception:
            self.log("[BŁĄD] Eksport bazy GDOŚ:\n" + traceback.format_exc())
            self.update_status("Błąd eksportu", "#D83B01", animate=False)
            return {"ok": False, "error": "Nie udało się wyeksportować (szczegóły w logu)"}

    def gdos_import_task(self):
        """Import bazy GDOŚ z pliku Excel wskazanego w polu."""
        try:
            we = getattr(self, "gdos_xlsx_entry", None)
            raw = we.get().strip() if we is not None else ""
            if not raw or not Path(raw).exists():
                self.log("[BŁĄD] Wskaż najpierw istniejący plik Excel do importu.")
                self.update_status("Brak pliku", "#D83B01", animate=False)
                return {"ok": False, "error": "Brak pliku"}
            n = self.gdos_importuj_excel(Path(raw))
            self.log(f"[OK] Zaimportowano bazę GDOŚ ({n} obszarów) z pliku: {raw}")
            self.update_status("Baza zaimportowana", "#107C10", animate=False)
            return {"ok": True, "count": n}
        except Exception:
            self.log("[BŁĄD] Import bazy GDOŚ:\n" + traceback.format_exc())
            self.update_status("Błąd importu", "#D83B01", animate=False)
            return {"ok": False, "error": "Nie udało się zaimportować (szczegóły w logu)"}

    def notify_update_check(self, is_latest, manual=False):
        """Wynik AUTOMATYCZNEGO sprawdzenia wersji przy starcie —
        dyskretne potwierdzenie, żeby nie wyglądało jakby program nie sprawdzał."""
        if is_latest and not manual:
            self._emit({"type": "toast", "kind": "ok",
                        "text": "Program jest w najnowszej wersji (" + CURRENT_VERSION + ")"})

    def check_update(self):
        threading.Thread(target=self.check_github_update, kwargs={"manual": True},
                         daemon=True).start()
        return {"ok": True}

    # ------------------------------------------------------------ polling

    def poll(self):
        with self._ev_lock:
            events = list(self._events)
            self._events.clear()
        return events or []

    def get_config(self):
        """Pełna konfiguracja startowa dla frontendu."""
        values = {}
        for tab in self.schema["tabs"]:
            for c in tab["controls"]:
                cid = c.get("id")
                if not cid:
                    continue
                kind = c["kind"]
                if kind in ("path", "text"):
                    saved = self.get_setting(f"web.{cid}", None)
                    values[cid] = saved if saved is not None else (c.get("default", "") or "")
                elif kind == "check":
                    saved = self.get_setting(f"web.{cid}", None)
                    values[cid] = bool(c.get("default", False)) if saved is None else bool(saved)
                elif kind == "select":
                    values[cid] = c.get("default", "")
                elif kind == "checks":
                    values[cid] = {ch: (ch == "Wszystkie") for ch in c["choices"]}
                elif kind == "margins":
                    saved = load_margins().get(c["mode"], {})
                    values[cid] = {
                        ftype: {s: str(saved.get(ftype, {}).get(s, d))
                                for s, d in (("T", "1.5"), ("B", "1.5"),
                                             ("L", "2.5"), ("R", "1.5"))}
                        for ftype in MARGIN_FILE_TYPES}
                elif kind == "fonts":
                    values[cid] = {f["sheet"]: str(self.get_setting(
                        f"web.font.{f['sheet']}", f["default"])) for f in c["fonts"]}
        return {"schema": self.schema, "values": values,
                "running": self.running}

    # ---------------------------------------------- układ PDF (kolejność)

    def get_pdf_order(self, mode):
        dst = None
        ent = self.entries.get(mode, {}).get("dst")
        if ent is not None:
            dst = ent.get().strip()
        if not dst:
            return {"ok": False, "error": "Wskaż najpierw lokalizację docelową."}
        config_folder = Path(dst) / "PDF"
        config_folder.mkdir(parents=True, exist_ok=True)
        order = get_saved_template_order(config_folder, mode)
        excluded = get_saved_excluded_templates(config_folder, mode)
        # od v2.0.32: 'Opis ogólny' jest stałą częścią zestawienia (1-Click go
        # generuje) — gdyby był wykluczony, przywracamy go na pozycję za stroną tytułową
        if "OPIS" in excluded:
            excluded = [k for k in excluded if k != "OPIS"]
            self.log("[UKŁAD] 'Opis ogólny' był wykluczony — przywrócono go "
                     "do scalania (zaraz za stroną tytułową).")
            set_saved_excluded_templates(config_folder, mode, excluded)
            set_saved_template_order(config_folder, mode, order)
        return {"ok": True, "order": order, "excluded": excluded,
                "dst": str(config_folder)}

    def save_pdf_order(self, mode, order, excluded=None):
        dst = None
        ent = self.entries.get(mode, {}).get("dst")
        if ent is not None:
            dst = ent.get().strip()
        if not dst:
            return {"ok": False, "error": "Wskaż najpierw lokalizację docelową."}
        config_folder = Path(dst) / "PDF"
        config_folder.mkdir(parents=True, exist_ok=True)
        set_saved_template_order(config_folder, mode, list(order or []))
        set_saved_excluded_templates(config_folder, mode, list(excluded or []))
        msg = f"[UKŁAD] Zapisano kolejność PDF dla trybu {mode}."
        if excluded:
            _lbl = ", ".join(t["label"] for t in PDF_ORDER_TEMPLATES
                             if t["key"] in excluded)
            msg += f" Wykluczono: {_lbl}."
        self.log(msg)
        return {"ok": True}

    # ------------------------------------------- ręczne scalanie PDF (web)

    def manual_merge_list(self):
        src = self.manual_pdf_src.get().strip() if self.manual_pdf_src else ""
        if not src or not Path(src).exists():
            return {"ok": False, "error": "Wskaż folder z plikami PDF."}
        pdfs = sorted(p.name for p in Path(src).glob("*.pdf"))
        return {"ok": True, "files": pdfs, "src": src}

    def manual_merge(self, order, filename="PDF_polaczony.pdf"):
        src = self.manual_pdf_src.get().strip() if self.manual_pdf_src else ""
        dst = self.manual_pdf_dst.get().strip() if self.manual_pdf_dst else ""
        if not src or not Path(src).exists():
            return {"ok": False, "error": "Wskaż folder z plikami PDF."}
        if not dst:
            return {"ok": False, "error": "Wskaż folder docelowy."}
        if not order:
            return {"ok": False, "error": "Lista plików jest pusta."}
        out = Path(dst)
        out.mkdir(parents=True, exist_ok=True)
        out_file = out / (filename or "PDF_polaczony.pdf")
        try:
            from pypdf import PdfWriter
            writer = PdfWriter()
            self.last_output_dir = out
            # spis treści = zakładki (outline) PDF z pełnymi tytułami dokumentów
            from app.config import PDF_ORDER_TEMPLATES, template_matches
            from pypdf import PdfReader
            current_page = 0
            for name in order:
                p = Path(src) / name
                if p.exists():
                    friendly_name = p.stem
                    for tpl in PDF_ORDER_TEMPLATES:
                        if template_matches(tpl, p.name):
                            friendly_name = tpl["label"]
                            break
                    try:
                        n_stron = len(PdfReader(str(p)).pages)
                    except Exception:
                        n_stron = 0
                    writer.add_outline_item(friendly_name, current_page)
                    current_page += n_stron
                    writer.append(str(p))
                else:
                    self.log(f"[SCAL] Pomijam (brak): {name}")
            with open(out_file, "wb") as f:
                writer.write(f)
            self.log(f"[SCAL] Zapisano: {out_file}")
            return {"ok": True, "file": str(out_file)}
        except Exception as e:
            self.log(f"[BŁĄD] Scalanie: {e}\n{traceback.format_exc()}")
            return {"ok": False, "error": str(e)}

    # --------------------------------- worker Word (z poprawnym wejściem)

    def task_word_processing_subprocess(self, in_dir, out_dir, remove_names,
                                        file_filter=None, margins_dict=None):
        """Kopia z tab_word, ale worker szuka main.py LUB main_web.py."""
        try:
            from app.gui.tabs.tab_word import normalize_filter_selection
        except Exception:
            normalize_filter_selection = lambda f: list(f or [])

        python_exe = sys.executable
        frozen = bool(getattr(sys, "frozen", False))
        if frozen:
            # EXE (PyInstaller): worker obsługuje ten sam plik wykonywalny
            worker_script = None
        else:
            here = Path(__file__).resolve()
            root = here.parents[2]
            worker_script = next((str(p) for p in (root / "main.py", root / "main_web.py")
                                  if p.exists()), str(root / "main.py"))
        remove_flag = " --remove-names" if remove_names else ""
        selected_filters = normalize_filter_selection(file_filter)
        filter_flag = ("" if "WSZYSTKIE" in selected_filters else
                       "".join(f' --filter "{flt}"' for flt in sorted(selected_filters)))

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
            cmd_base = (f'"{python_exe}" --word-worker '
                        f'"{str(in_dir).rstrip("/")}" "{str(out_dir).rstrip("/")}" '
                        f'--log-file "{log_path}"')
        else:
            cmd_base = (f'"{python_exe}" -u "{worker_script}" --word-worker '
                        f'"{str(in_dir).rstrip("/")}" "{str(out_dir).rstrip("/")}" '
                        f'--log-file "{log_path}"')
        bat_content = (f"@echo off\nchcp 65001 >nul\n"
                       f"set PYTHONIOENCODING=utf-8\nset PYTHONUNBUFFERED=1\n"
                       f"{cmd_base}{remove_flag}{filter_flag}{margins_flag}\n"
                       f"exit /b %errorlevel%\n")
        with tempfile.NamedTemporaryFile("w", suffix=".bat", delete=False,
                                         encoding="utf-8") as bat_file:
            bat_file.write(bat_content)
        bat_path = bat_file.name
        process = None
        try:
            process = subprocess.Popen(["cmd", "/c", bat_path],
                                        creationflags=subprocess.CREATE_NO_WINDOW)
            with open(log_path, "r", encoding="utf-8") as f:
                while True:
                    if self.stop_event.is_set():
                        try:
                            subprocess.run(["taskkill", "/F", "/T", "/PID", str(process.pid)],
                                           stdout=subprocess.DEVNULL,
                                           stderr=subprocess.DEVNULL)
                        except Exception:
                            process.kill()
                        self.log("Proces ukrytego Worda (MIETEK) zablokowany "
                                 "i ugaszony z powodzeniem.")
                        raise InterruptedError()
                    line = f.readline()
                    if line:
                        self.log(line.rstrip())
                    elif process.poll() is not None:
                        for remaining in f.readlines():
                            if remaining:
                                self.log(remaining.rstrip())
                        break
                    else:
                        time.sleep(0.1)
            if process.poll() not in (0, None):
                self.log(f"[WORD] Proces Workera zakończył się kodem {process.poll()} "
                         "— sprawdź log powyżej.")
        finally:
            if process is not None and process.poll() is None:
                try:
                    subprocess.run(["taskkill", "/F", "/T", "/PID", str(process.pid)],
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL)
                except Exception:
                    process.kill()
            for p in (bat_path, log_path, margins_file_path):
                if p:
                    try:
                        os.remove(p)
                    except Exception:
                        pass

    # ---------------------------------------- podgląd "na żywo" (live stream)

    def init_live_stream(self, total_files=None):
        """Odpowiednik live-streamu z CTk — w web GUI trafia do dziennika."""
        self.stream_queue = []
        self.stream_current = None
        self.stream_completed = []
        self.stream_start_time = time.time()
        self.stream_files_count = 0
        if total_files:
            self.log(f"[PRZYGOTOWANO] {total_files} plików do przetworzenia...")

    def add_to_stream_queue(self, source_path, target_path=None):
        self.stream_queue.append({
            "source": str(source_path),
            "target": str(target_path) if target_path else "...",
            "status": "pending",
        })

    def start_stream_file(self, source_path, target_path=None):
        from pathlib import Path as _P
        name = _P(str(source_path)).name
        self.stream_current = {
            "source": str(source_path),
            "target": str(target_path) if target_path else "...",
            "start_time": time.time(),
            "status": "processing",
        }
        self.stream_queue = [q for q in self.stream_queue
                             if q["source"] != str(source_path)]
        self.log(f"→ przetwarzam: {name}")

    def complete_stream_file(self, source_path, target_path, duration=None):
        from pathlib import Path as _P
        if duration is None and self.stream_current:
            duration = time.time() - self.stream_current["start_time"]
        self.stream_completed.append({
            "source": str(source_path),
            "target": str(target_path),
            "duration": duration,
            "status": "completed",
        })
        self.stream_files_count += 1
        self.stream_current = None
        if len(self.stream_completed) > 15:
            self.stream_completed = self.stream_completed[-15:]
        name = _P(str(source_path)).name
        secs = f"{duration:.1f}s" if duration is not None else "?"
        self.log(f"✓ gotowe: {name} ({secs})")

    # ---------------------------------------------- okno walidacji (web)

    def show_validation_window_sync(self, title_text, warnings):
        """Zastępuje ValidationWindow: dialog w przeglądarce, czeka na decyzję."""
        lines = "\n".join(f"• {w}" for w in (warnings or []))
        message = (f"{title_text}\n\n{lines}\n\n"
                   "Czy kontynuować mimo ostrzeżeń?")
        return self._web_confirm(title_text, message)

    def _web_confirm(self, title, message):
        did = f"v{time.time():.6f}".replace(".", "")
        ev = threading.Event()
        self._dialog_waits[did] = (ev, [False])
        self._emit({"type": "dialog", "kind": "confirm", "id": did,
                    "title": str(title), "message": str(message)})
        ev.wait(timeout=3600)
        return bool(self._dialog_waits.pop(did, (None, [False]))[1][0])

    # ------------------------------------ wybór plików/folderu z pipeline'ów

    def select_dir(self, entry_widget):
        path = self._dialog_path("folder")
        if path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)
        return path

    def select_file(self, entry_widget, filetypes=None):
        path = self._dialog_path("file")
        if path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)
        return path

    def select_save_file(self, entry_widget):
        path = self._dialog_path("save")
        if path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)
        return path

    def _dialog_path(self, kind):
        try:
            import webview
            win = webview.windows[0]
            try:
                _fd = webview.FileDialog
                d_folder, d_open, d_save = _fd.FOLDER, _fd.OPEN, _fd.SAVE
            except Exception:
                d_folder = getattr(webview, "FOLDER_DIALOG", None)
                d_open = getattr(webview, "OPEN_DIALOG", None)
                d_save = getattr(webview, "SAVE_DIALOG", None)
            if kind == "folder":
                res = win.create_file_dialog(d_folder)
                p = res[0] if res else None
            elif kind == "save":
                res = win.create_file_dialog(d_save, save_filename="Dokument.docx")
                p = res if isinstance(res, str) else (res[0] if res else None)
            else:
                res = win.create_file_dialog(d_open)
                p = res[0] if res else None
        except Exception as e:
            self.log(f"[BŁĄD] Wybór pliku: {e}")
            return None
        if p:
            self.add_to_history(str(p) if kind == "folder" else str(Path(p).parent))
        return p

    # ------------------------------------------------------- zamykanie

    def destroy(self):
        try:
            import webview
            for w in webview.windows:
                w.destroy()
        except Exception:
            pass
