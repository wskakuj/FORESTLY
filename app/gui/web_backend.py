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

from app.config import (
    load_order_store,CURRENT_VERSION, SETTINGS_FILE, HISTORY_FILE,
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
    "gdos_xlsx": ("Arkusz Excel", ("*.xlsx", "*.xls")),
    "all_skroty": ("Word i PDF", ("*.docx", "*.doc", "*.pdf")),
    "mt_template": ("Dokument Word", ("*.docx", "*.doc")),
    "tt_template": ("Dokument Word", ("*.docx", "*.doc")),
    "zm_src": ("Baza Access", ("*.mdb",)),
}
MARGIN_FILE_TYPES = ["REJESTR1", "OPTAX", "TAB_KLW3", "WSKAZ1", "HALIZNY",
                     "WYK_NEG", "OPIS", "ZEST1", "WK_ZM1", "SKROTY"]

# domyślne marginesy różne od standardowych (1.5/1.5/2.5/1.5) — SKROTY
# zachowuje dzisiejszy wygląd wykazu skrótów
MARGIN_DOMYSLNE_TYPY = {"SKROTY": {"T": "1.3", "B": "1.5", "L": "1.1", "R": "1.1"}}


# --------------------------------------------------------------- sztuczne widgety

# ------------------------------------------------------------ okno podglądu
# Osobne okno systemowe z podglądem marginesów/czcionek: można je przesunąć
# w dowolne miejsce EKRANU (nie tylko wewnątrz programu) i dowolnie skalować.
# Zawartość sama się odświeża — okno czyta aktualne marginesy i czcionki
# zapisane w ustawieniach (te same, które edytuje kreator).
_PV_OKNO_HTML = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  html, body { height: 100%; }
  body { margin: 0; background: #23262d; color: #e8e8e8;
         font: 13px "Segoe UI", sans-serif; display: flex;
         flex-direction: column; overflow: hidden; }
  header { display: flex; gap: 12px; align-items: center;
           padding: 10px 14px 8px; }
  header strong { font-size: 14px; }
  select { background: #333; color: #eee; border: 1px solid #555;
           border-radius: 6px; padding: 5px 9px; font: inherit; }
  #drukuj { background: #33553f; color: #e8f0ea; border: 1px solid #567a63;
            border-radius: 6px; padding: 5px 12px; font: inherit;
            cursor: pointer; }
  #drukuj:hover { background: #3f6850; }
  #info { color: #98a0aa; font-size: 12px; white-space: nowrap;
          overflow: hidden; text-overflow: ellipsis; flex: 1; }
  #wrap { flex: 1; overflow: auto; padding: 8px 16px 16px; }
  #sheet { position: relative; margin: 0 auto; background: #fff;
           box-shadow: 0 6px 26px rgba(0, 0, 0, .6); }
  iframe { position: absolute; inset: 0; border: 0; width: 100%; height: 100%; }
</style></head><body>
<header>
  <strong>Podgląd</strong>
  <select id="typ"></select>
  <button id="drukuj" title="Wydrukuj jedną stronę testową — na aktualnych czcionkach i marginesach">Drukuj</button>
  <span id="info"></span>
</header>
<div id="wrap"><div id="sheet"><iframe id="frame"></iframe></div></div>
<script>
(function () {
  const TYPY = ["OPTAX", "REJESTR1", "WSKAZ1", "TAB_KLW3", "WSK_ZB",
                "ZEST1", "HALIZNY", "WYK_NEG", "WK_ZM1", "SKROTY"];
  const MP_WIERSZ = { REJESTR1: "REJESTR1", TAB_KLW3: "TAB_KLW3" };
  let typ = "OPTAX", lastHtml = "", poziom = false, ileStron = 1;
  const sel = document.getElementById("typ");
  TYPY.forEach(t => {
    const o = document.createElement("option");
    o.value = t; o.textContent = t; sel.appendChild(o);
  });
  sel.onchange = () => { typ = sel.value; lastHtml = ""; odswiez(); };

  function dopasuj() {
    const szer = poziom ? 1123 : 794, wys = poziom ? 794 : 1123;
    const w = document.getElementById("wrap");
    /* wiele stron (SKROTY): kartka dopasowana szerokością, resztę
       przewija się w dół — widać, co na którą stronę się przesunęło */
    const total = wys * ileStron + (ileStron > 1 ? 18 * (ileStron - 1) : 0);
    const sk = ileStron > 1
      ? Math.max(.1, Math.min((w.clientWidth - 32) / szer, 1))
      : Math.max(.1, Math.min((w.clientWidth - 32) / szer,
                              (w.clientHeight - 24) / wys));
    const sh = document.getElementById("sheet");
    sh.style.width = szer + "px"; sh.style.height = total + "px";
    sh.style.transform = "scale(" + sk + ")";
    sh.style.transformOrigin = "top left";
    sh.style.marginBottom = (total * (sk - 1)) + "px";
  }
  document.getElementById("info").textContent = "Ładowanie podglądu…";
  async function odswiez() {
    try {
      if (!window.pywebview || !pywebview.api) {
        setTimeout(odswiez, 250);   /* api wstrzykiwane asynchronicznie */
        return;
      }
      const r = await pywebview.api.preview_window_state(typ);
      if (!r || !r.ok) {
        document.getElementById("info").textContent =
          (r && r.error) || "Nie udało się przygotować podglądu.";
        return;
      }
      document.getElementById("info").textContent =
        (r.zrodlo || "dokument przykładowy") + "  (" + typ +
        (r.poziom ? ", poziomo" : ", pionowo") + ")";
      if (r.html !== lastHtml) {
        lastHtml = r.html;
        poziom = !!r.poziom;
        const m = r.marginesy || {};
        const T = m.T != null ? m.T : 1.5, B = m.B != null ? m.B : 1.5,
              L = m.L != null ? m.L : 2.5, R = m.R != null ? m.R : 1.5;
        const szer = poziom ? 1123 : 794, wys = poziom ? 794 : 1123;
        const css = "<style>html, body { background:#fff !important; " +
          "max-width:none !important; margin:0 !important; padding:0 " +
          "important; overflow:hidden !important; } " +
          ".mp-page { position:relative; width:" + szer + "px; height:" +
          wys + "px; margin:0 0 18px; } " +
          ".mp-page:last-child { margin-bottom:0; } " +
          ".mp-win { position:absolute; left:" + L + "cm; top:" + T + "cm; " +
          "right:" + R + "cm; bottom:" + B + "cm; overflow:hidden; } " +
          "</style>" +
          "<scr" + "ipt>(function () { " +
          "function mpStrona(dzieci) { " +
          "var p = document.createElement('div'); p.className = 'mp-page'; " +
          "var w = document.createElement('div'); w.className = 'mp-win'; " +
          "for (var i = 0; i < dzieci.length; i++) w.appendChild(dzieci[i]); " +
          "p.appendChild(w); document.body.appendChild(p); } " +
          "function mpWrap() { " +
          "if (document.body.getAttribute('data-mp') === 'done') return; " +
          "document.body.setAttribute('data-mp', 'done'); " +
          "var wszystkie = Array.prototype.slice.call(document.body.children); " +
          "var nodes = wszystkie.filter(function (n) { " +
          "return n.tagName !== 'STYLE' && n.tagName !== 'SCRIPT'; }); " +
          "var pierwsza = -1; " +
          "for (var i = 0; i < nodes.length; i++) { " +
          "if (nodes[i].classList && nodes[i].classList.contains('sk-strona')) { " +
          "pierwsza = i; break; } } " +
          "if (pierwsza < 0) { mpStrona(nodes); return; } " +
          "mpStrona(nodes.slice(0, pierwsza + 1)); " +
          "for (var j = pierwsza + 1; j < nodes.length; j++) mpStrona([nodes[j]]); } " +
          "if (document.readyState === 'loading') " +
          "document.addEventListener('DOMContentLoaded', mpWrap); " +
          "else mpWrap(); })();</scr" + "ipt>";
        const fr = document.getElementById("frame");
        fr.onload = function () {
          try {
            const n = Math.max(1,
              fr.contentDocument.querySelectorAll(".mp-page").length);
            if (n !== ileStron) { ileStron = n; dopasuj(); }
          } catch (e) { /* iframe niedostępny */ }
        };
        fr.srcdoc = r.html + css;
      }
      dopasuj();
    } catch (e) { /* okno się zamyka */ }
  }

  /* wydruk JEDNEJ strony testowej — dokładnie te czcionki i marginesy,
     które widać w podglądzie (pierwsza kartka) */
  document.getElementById("drukuj").onclick = function () {
    var fr = document.getElementById("frame");
    try {
      var doc = fr.contentDocument;
      if (!doc || !doc.body || !doc.querySelector(".mp-page")) return;
      var st = doc.getElementById("mp-print");
      if (!st) {
        st = doc.createElement("style"); st.id = "mp-print";
        doc.head.appendChild(st);
      }
      st.textContent = "@page { size: A4 " + (poziom ? "landscape" : "portrait")
                     + "; margin: 0 }"
                     + " @media print { .mp-page + .mp-page { display: none } }";
      fr.contentWindow.print();
    } catch (e) { /* okno się zamyka */ }
  };

  window.odswiezNatychmiast = () => { lastHtml = ""; odswiez(); };
  window.addEventListener("resize", dopasuj);
  setInterval(odswiez, 900);
  odswiez();
})();
</scr""" + """ipt></body></html>"""


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


def _marginesy_z_slownika(saved):
    """{T,B,L,R} (z dysku / frontendu) -> [góra, dół, lewo, prawo] (float).

    Takiej listy oczekuje szablony.py (konwersja jak w torze 1-Click);
    błędne wpisy są po prostu pomijane.
    """
    out = {}
    for ftype, poz in (saved or {}).items():
        try:
            out[ftype] = [
                float(str(poz.get("T", "1.5")).replace(",", ".")),
                float(str(poz.get("B", "1.5")).replace(",", ".")),
                float(str(poz.get("L", "2.5")).replace(",", ".")),
                float(str(poz.get("R", "1.5")).replace(",", "."))]
        except (TypeError, ValueError, AttributeError):
            continue
    return out


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

        # folder wyników zapamiętany z poprzedniego uruchomienia (settings.json)
        try:
            _d = (self.load_settings() or {}).get("last_output_dir")
            if _d and Path(_d).exists():
                self.last_output_dir = Path(_d)
        except Exception:
            pass

        # zdarzenia do frontendu + dialogi
        self._events = deque()
        self._ev_lock = threading.Lock()
        self._dialog_waits = {}   # id -> (Event, result)
        self._pc_pliki = []       # pliki przeciągnięte do konwertera PDF

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

        # remove_names_var potrzebny dla ALL/WORD — JEDNA wspólna wartość.
        # Uwaga: zakładka Word ma własny checkbox na tej samej zmiennej
        # (word_remove_names) i wcześniej NADPISYWAŁ ją swoim domyślnym
        # False — dlatego zawsze przywracamy stan z ustawienia 1-Click.
        self.remove_names_var = FakeVar(
            bool(self.get_setting("web.remove_names", True)))

        # nowe szablony wydruków (HTML → PDF, bez Worda) — przełącznik
        # na pierwszym ekranie kreatora 1-Click; stan pamiętany w ustawieniach
        self.nowe_szablony_var = FakeVar(
            bool(self.get_setting("web.nowe_szablony", False)))


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
                # 'Nowe Szablony': przy pierwszym otwarciu pokaż wartości
                # kreatora Pełnego Automatu, żeby tabela nie kłamała
                if not saved and mode == "NS":
                    saved = load_margins().get("ALL", {})
                for ftype in MARGIN_FILE_TYPES:
                    fsaved = saved.get(ftype, {})
                    for side, dflt in (("T", "1.5"), ("B", "1.5"),
                                       ("L", "2.5"), ("R", "1.5")):
                        dflt = (MARGIN_DOMYSLNE_TYPY.get(ftype) or {}).get(side, dflt)
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
        values = dict(values or {})

        # przełącznik "nowe szablony" z ekranu powitalnego kreatora 1-Click
        # (kontrolka spoza schematu — własny HTML w kreatorze)
        if "nowe_szablony" in values:
            on = bool(values.pop("nowe_szablony"))
            self.nowe_szablony_var.set(on)
            try:
                self.set_setting("web.nowe_szablony", on)
            except Exception:
                pass

        controls = {}

        def _index(ctrls):
            # schodzimy też do grup (zwijane sekcje schematu) — kontrolki
            # w nich (daty kreatora 1-Click, woj./pow./gmina itd.) muszą
            # być indeksowane jak każde inne; wcześniej lądowały po cichu
            # w koszu i backend pracował na wartościach DOMYŚLNYCH
            for c in ctrls:
                if c.get("kind") == "group":
                    _index(c.get("controls") or [])
                elif c.get("id"):
                    controls[c["id"]] = c

        for tab in self.schema["tabs"]:
            _index(tab["controls"])
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
            elif kind == "czcionki":
                # czcionki raportów nowego wyglądu (tytuł / tabela) — zapis
                # do settings.json, z walidacją rozmiarów w trakcie pisania
                mode_cfg = {}
                all_valid = True
                for typ_, sekcje in (val or {}).items():
                    mode_cfg[typ_] = {}
                    for gdzie in ("tytul", "tabela"):
                        c_ = (sekcje or {}).get(gdzie) or {}
                        try:
                            pt = float(str(c_.get("pt", "")).replace(",", "."))
                        except (TypeError, ValueError):
                            all_valid = False
                            pt = None
                        mode_cfg[typ_][gdzie] = {
                            "pt": pt if pt is not None else
                                  (12.0 if gdzie == "tytul" else 8.6),
                            "font": str(c_.get("font") or "")[:40]}
                if mode_cfg and all_valid:
                    self.set_setting(f"web.czcionki.{c['mode']}", mode_cfg)
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
        # przy "obu wersjach" pasek statusu mówi, który przebieg trwa
        et = getattr(self, "_przebieg_etykieta", "")
        if et:
            text = f"{et} · {text}"
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
            "start_nowe_szablony": self.start_nowe_szablony,
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
        # po chwili ubij procesy Office zostawione w tle — zwalniają
        # blokady na folderach wynikowych (patrz app/core/office_guard)
        threading.Thread(target=self._office_sweep_after_stop,
                         daemon=True).start()
        return {"ok": True}

    def _office_sweep_after_stop(self):
        time.sleep(2.5)
        try:
            from app.core import office_guard
            office_guard.kill_registered(log=self.log)
        except Exception:
            pass

    def _zapamietaj_folder_wynikow(self, path):
        """Zapisuje folder wyników w settings.json — przetrwa restart GUI
        (przycisk 'Otwórz folder wyników' działa też po ponownym uruchomieniu)."""
        try:
            ustaw = self.load_settings() or {}
            ustaw["last_output_dir"] = str(path)
            self.save_settings(ustaw)
        except Exception:
            pass

    def open_last_output(self):
        """Otwiera Eksplorator w folderze z finalnymi wynikami ostatniego zadania.

        Dla Pełnego Automatu otwieramy od razu folder 'PDF polaczone'
        (finalne pliki), a nie folder nadrzędny z całą strukturą etapów.
        """
        try:
            d = getattr(self, "last_output_dir", None)
            if not (d and Path(d).exists()):
                # awaryjnie: folder zapisany w ustawieniach (działa po restarcie)
                try:
                    d = (self.load_settings() or {}).get("last_output_dir") or d
                except Exception:
                    pass
            if not (d and Path(d).exists()):
                # awaryjnie: folder docelowy 1-Click z bieżących ustawień
                try:
                    e = (self.entries or {}).get("ALL", {}).get("dst")
                    alt = e.get() if e is not None else ""
                    if alt and Path(alt).exists():
                        d = alt
                except Exception:
                    d = None
            if not (d and Path(d).exists()):
                return {"ok": False,
                        "error": "Folder wyników nieznany \u2014 uruchom najpierw zadanie."}
            target = Path(d)
            final = target / "PDF polaczone"
            if final.exists():
                target = final
            if os.name == "nt":
                try:
                    os.startfile(str(target))
                except Exception:
                    # awaryjnie przez explorer.exe (inne błędy ShellExecute)
                    subprocess.Popen(["explorer", str(target)])
            else:
                self.log("[UWAGA] Otwieranie folderu jest dostępne na Windows.")
            return {"ok": True, "path": str(target)}
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
            self.log(f"[OK] Zaimportowano bazę GDOŚ ({n} obszarów) z pliku: {raw} "
                     "— poprzednia baza została w całości zastąpiona.")
            self.update_status("Baza zaimportowana", "#107C10", animate=False)
            # edytor (jeśli otwarty) musi przeładować dane, żeby nie nadpisał
            # nowej bazy starym stanem z pamięci
            self._emit({"type": "gdos_changed", "count": n})
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

    # --------------------------------------- przeciąganie plików (drag&drop)
    def set_pv_window(self, window):
        """Zapamiętuje okno pywebview i włącza natywne przeciąganie —
        dzięki temu drop z Eksploratora daje PRAWDZIWE ścieżki (WebView2),
        także folderów; pobiera je drop_paths()."""
        self._pv_window = window
        try:
            window.dom.get_element("body").events.drop += self._natywny_drop
        except Exception as e:
            self.log(f"[DnD] Nie udało się włączyć przeciągania plików: {e}")

    def _natywny_drop(self, *args, **kwargs):
        pass    # ścieżki zbiera pywebview — pobiera je dopiero drop_paths()

    def drop_paths(self, kind="file"):
        """Ścieżki z ostatniego przeciągnięcia (pliki albo folder).
        Frontend woła to zaraz po zdarzeniu drop."""
        try:
            from webview.dom import _dnd_state
        except Exception:
            return {"ok": False, "error": "no_dnd"}
        paths = [pp[1] for pp in list(_dnd_state.get("paths", []))]
        _dnd_state["paths"] = []
        if not paths:
            return {"ok": False, "error": "empty"}
        if str(kind).lower() == "folder":
            foldery = [pp for pp in paths if Path(pp).is_dir()]
            if not foldery:
                return {"ok": False, "error": "not_folder", "paths": paths}
            return {"ok": True, "path": foldery[0], "paths": paths}
        pliki = [pp for pp in paths if Path(pp).is_file()]
        if not pliki:
            return {"ok": False, "error": "not_file", "paths": paths}
        return {"ok": True, "path": pliki[0], "paths": paths}

    # --------------------------------------- konwerter PDF: pliki z drop
    def pdfconv_drop_add(self, paths):
        """Dopisuje przeciągnięte pliki do kolejki konwertera PDF."""
        dodane = 0
        for x in (paths or []):
            try:
                if Path(x).is_file() and x not in self._pc_pliki:
                    self._pc_pliki.append(x)
                    dodane += 1
            except Exception:
                pass
        self._emit({"type": "pdfconv_files", "files": list(self._pc_pliki)})
        return {"ok": True, "count": dodane, "files": list(self._pc_pliki)}

    def pdfconv_drop_clear(self):
        self._pc_pliki = []
        self._emit({"type": "pdfconv_files", "files": []})
        return {"ok": True}

    def check_update(self):
        threading.Thread(target=self.check_github_update, kwargs={"manual": True},
                         daemon=True).start()
        return {"ok": True}

    # --------------------------------- aktualizator: wygląd jak cały program
    def updater_ask(self, title, message):
        """Pytanie aktualizatora w modalu programu (nie tkinter!)."""
        return self._web_confirm(title, message)

    def updater_info(self, title, message):
        self._emit({"type": "dialog", "kind": "info", "title": str(title),
                    "message": str(message)})

    def updater_error(self, title, message):
        self._emit({"type": "dialog", "kind": "error", "title": str(title),
                    "message": str(message)})

    def mapa_drop_clear(self):
        """Zapomina mapy przeciągnięte w kreatorze (folder tymczasowy)."""
        import shutil
        folder = Path(tempfile.gettempdir()) / "forestly_mapy"
        try:
            if folder.is_dir():
                shutil.rmtree(folder)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ------------------------------------------------------------ polling

    def poll(self):
        with self._ev_lock:
            events = list(self._events)
            self._events.clear()
        return events or []

    def get_config(self):
        """Pełna konfiguracja startowa dla frontendu."""
        values = {}

        def _flat(ctrls):
            # wraz z grupami (zwijane sekcje) — inaczej wartości pól
            # zagnieżdżonych nigdy nie trafiały do frontendu
            out = []
            for c in ctrls:
                if c.get("kind") == "group":
                    out.extend(_flat(c.get("controls") or []))
                elif c.get("id"):
                    out.append(c)
            return out

        for tab in self.schema["tabs"]:
            for c in _flat(tab["controls"]):
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
                    if not saved and c["mode"] == "NS":
                        saved = load_margins().get("ALL", {})
                    values[cid] = {
                        ftype: {s: str(saved.get(ftype, {}).get(
                                    s, (MARGIN_DOMYSLNE_TYPY.get(ftype) or {}).get(s, d)))
                                for s, d in (("T", "1.5"), ("B", "1.5"),
                                             ("L", "2.5"), ("R", "1.5"))}
                        for ftype in MARGIN_FILE_TYPES}
                elif kind == "czcionki":
                    saved = self.get_setting(
                        f"web.czcionki.{c['mode']}", None)
                    if not isinstance(saved, dict) and c["mode"] == "NS":
                        saved = self.get_setting("web.czcionki.ALL", None)
                    if not isinstance(saved, dict):
                        saved = {}
                    defcz = {"tytul": {"pt": "12", "font": ""},
                             "tabela": {"pt": "8.6", "font": ""}}
                    values[cid] = {
                        t: {g: dict(saved.get(t, {}).get(g)
                                   or defcz[g])
                            for g in ("tytul", "tabela")}
                        for t in (c.get("types") or [])}
                elif kind == "fonts":
                    values[cid] = {f["sheet"]: str(self.get_setting(
                        f"web.font.{f['sheet']}", f["default"])) for f in c["fonts"]}
        return {"schema": self.schema, "values": values,
                "running": self.running}

    # ------------------------------------------------ nowe szablony (pojedynczy raport)

    def start_nowe_szablony(self):
        """NOWE SZABLONY: jeden typ raportu TXT → HTML → PDF (bez Worda).

        Skrót do nowych szablonów z Pełnego Automatu — np. do dogenerowania
        samych WYK_NEG albo sprawdzenia wyglądu jednego raportu.
        """
        from tkinter import messagebox
        src = (self.entries.get("NS", {}).get("src") or FakeEntry("")).get().strip()
        dst = (self.entries.get("NS", {}).get("dst") or FakeEntry("")).get().strip()
        typ = str(self.ns_typ_var.get() or "WYK_NEG").strip().upper()
        bez_nazwisk = bool(self.ns_bez_nazwisk_var.get())
        if not src or not Path(src).exists():
            messagebox.showwarning("Błąd", "Wybierz istniejący folder z Mietkami (obręby).")
            return
        if not dst:
            messagebox.showwarning("Błąd", "Wskaż folder docelowy (PDF).")
            return
        if self.running:
            return
        self.last_output_dir = Path(dst)
        self._disable_ui_for_process()
        self.log(f"[NOWE SZABLONY] {typ} — Mietki: {src} → PDF: {dst}"
                 + (" (bez nazwisk)" if bez_nazwisk else ""))
        threading.Thread(target=self._nowe_szablony_thread,
                         args=(src, dst, typ, bez_nazwisk), daemon=True).start()

    def _nowe_szablony_thread(self, src, dst, typ, bez_nazwisk):
        try:
            from app.core import szablony
            from app.core.wydruki import (generuj_wszystkie_po_przeniesieniu,
                                           generuj_halizny_txt, czytaj_agencje)
            from app.gui.tabs.tab_wydruki import AGENCJA_NAGLOWKA
            if typ not in szablony.RENDERERY:
                raise RuntimeError(f"Nieznany typ raportu: {typ}")
            self.update_status(f"Nowe szablony: {typ} — TXT z DBF, potem PDF...",
                               "#0078D7")
            src_p, dst_p = Path(src), Path(dst)
            obraby = sorted([d for d in src_p.iterdir() if d.is_dir()])
            if not obraby:
                raise RuntimeError("Brak podfolderów obrębów we wskazanym folderze.")
            # marginesy: własne dla tej zakładki (tryb NS); gdy użytkownik
            # nie ruszał tabeli — przejmujemy zapamiętane z Pełnego Automatu
            for _tryb in ("NS", "ALL"):
                margins = _marginesy_z_slownika(load_margins().get(_tryb))
                if margins:
                    break
            # czcionki: własne dla zakładki, z fallbackiem na kreator
            czc = self.get_setting("web.czcionki.NS", None)
            if not isinstance(czc, dict) or not czc:
                czc = self.get_setting("web.czcionki.ALL", None)
            if not isinstance(czc, dict):
                czc = {}
            self.start_progress_tracking(len(obraby), f"Nowe szablony: {typ}")
            ok, blad, pominiete = 0, 0, []
            for i, obr in enumerate(obraby, 1):
                self.check_stop()
                try:
                    # 1) TXT z DBF-ów (tylko wybrany raport), tak jak w zakładce
                    #    'MIETEK -> TXT' — plik powstaje obok plików DBF obrębu
                    ag = czytaj_agencje(obr) or AGENCJA_NAGLOWKA
                    if typ == "HALIZNY":
                        try:
                            hp, _hn = generuj_halizny_txt(obr, agencja=ag or None)
                        except TypeError:   # starsza sygnatura bez 'agencja'
                            hp = generuj_halizny_txt(obr)
                        txt = Path(hp) if hp else None
                    else:
                        out = generuj_wszystkie_po_przeniesieniu(
                            obr, agencja=ag or None, tylko={f"{typ}.TXT"})
                        txt = Path(out[f"{typ}.TXT"]) if out.get(f"{typ}.TXT") else None
                    if txt is None or not txt.exists() or txt.stat().st_size < 100:
                        pominiete.append(obr.name)
                        self.log(f"  ℹ️ {obr.name}: {typ}.txt nie powstał (brak danych) — pomijam.")
                    else:
                        # 2) TXT → HTML → PDF nowym wyglądem
                        pdf = dst_p / obr.name / f"{typ}.pdf"
                        pdf.parent.mkdir(parents=True, exist_ok=True)
                        szablony.generuj_raport_pdf(
                            typ, txt, pdf,
                            bez_nazwisk=bool(bez_nazwisk and typ in szablony.USUWA_NAZWISKA),
                            margins=margins, czcionki=czc)
                        ok += 1
                        self.log(f"  ✅ {obr.name}: {typ}.txt + PDF → {pdf}")
                except Exception as e:
                    blad += 1
                    self.log(f"  ✗ {obr.name}: {e}")
                self.set_progress(i / len(obraby), current_file=obr.name)
            self.set_progress(1.0)
            self.last_output_dir = dst_p
            if blad:
                self.update_status(f"Zakończono z błędami ({ok} OK / {blad} błędów)",
                                   "#D83B01", animate=False)
            else:
                self.update_status(f"Gotowość: {ok} × {typ}.pdf", "#107C10", animate=False)
            self.log(f"[NOWE SZABLONY] Zrobione: {ok} PDF"
                     + (f", błędy: {blad}" if blad else "")
                     + (f", pominięte (brak danych): {', '.join(pominiete)}"
                        if pominiete else "")
                     + f". Folder: {dst}")
        except InterruptedError:
            self.update_status("Zatrzymano", "#D83B01", animate=False)
            self.log("[STOP] Przerwano generowanie nowych szablonów.")
        except Exception:
            self.log(traceback.format_exc())
            self.update_status("Błąd", "#D83B01", animate=False)
        finally:
            self.restore_all_buttons()

    # ------------------------------------------ podgląd marginesów (na żywo)

    def _find_preview_txt(self, typ):
        """TXT do podglądu: najpierw folder z formularza, potem cache z DBF.

        Nie ruszam folderów użytkownika — DBF-y obrębu są kopiowane do
        katalogu tymczasowego i tam powstaje TXT (cache na czas sesji).
        """
        import shutil
        from app.core.wydruki import (generuj_wszystkie_po_przeniesieniu,
                                       generuj_halizny_txt)
        from app.gui.tabs.tab_wydruki import AGENCJA_NAGLOWKA

        # 1) istniejące TXT (źródło/docelowy 1-Click i Nowych Szablonów);
        #    w trybie ALL pliki TXT z poprzedniego przebiegu leżą w dst/TXT
        kandydaci, widziane = [], set()
        for tryb in ("ALL", "NS"):
            for kt in ("src", "dst"):
                p = (self.entries.get(tryb, {}).get(kt)
                     or FakeEntry("")).get().strip()
                if p and Path(p).exists():
                    for d in (Path(p), Path(p) / "TXT"):
                        if d not in widziane:
                            widziane.add(d)
                            kandydaci.append(d)
        for d in kandydaci:
            for wz in (f"{typ}.TXT", f"{typ}.txt"):
                for f in sorted(d.rglob(wz)):
                    try:
                        if f.stat().st_size >= 100:
                            return f, str(d)
                    except OSError:
                        continue

        # 2) cache z DBF-ów (folder źródłowy 1-Click / NS)
        cache = getattr(self, "_preview_cache", None)
        if cache is None:
            cache = self._preview_cache = {}
        if typ in cache and cache[typ][0].exists():
            return cache[typ]
        for tryb in ("ALL", "NS"):
            p = (self.entries.get(tryb, {}).get("src")
                 or FakeEntry("")).get().strip()
            if not p or not Path(p).exists():
                continue
            obraby = sorted(d for d in Path(p).iterdir() if d.is_dir())
            for obr in obraby:
                if not (list(obr.glob("*.DBF")) or list(obr.glob("*.dbf"))):
                    continue
                work = Path(tempfile.gettempdir()) / "forestly_preview" / obr.name
                try:
                    if not work.exists():
                        shutil.copytree(obr, work)
                    if typ == "HALIZNY":
                        try:
                            hp, _hn = generuj_halizny_txt(work, agencja=AGENCJA_NAGLOWKA)
                        except TypeError:
                            hp = generuj_halizny_txt(work)
                        f = Path(hp) if hp else None
                    else:
                        out = generuj_wszystkie_po_przeniesieniu(
                            work, agencja=AGENCJA_NAGLOWKA, tylko={f"{typ}.TXT"})
                        f = Path(out[f"{typ}.TXT"]) if out.get(f"{typ}.TXT") else None
                except Exception:
                    continue
                if f and f.exists() and f.stat().st_size >= 100:
                    cache[typ] = (f, f"{obr.name} (z DBF)")
                    return cache[typ]
        return None, None

    def get_margins_preview(self, typ, margins, czcionki=None):
        """HTML podglądu raportu nowym wyglądem z podanymi marginesami
        i czcionkami (tytuł / tabela).

        Frontend pokazuje go w <iframe> i odświeża po każdej zmianie
        marginesu — bez generowania PDF, bez Worda.
        """
        from app.core import szablony
        typ = str(typ or "OPTAX").upper()
        if typ == "SKROTY":
            # wykaz skrótów i symboli — bez pliku TXT, z czcionkami z wiersza
            # SKROTY (marginesy ma zawsze domyślne)
            sk = None
            try:
                sk = self._resolve_skroty_path()
            except Exception:
                sk = None
            if not sk or not Path(sk).exists():
                return {"ok": False, "error":
                        "Nie znalazłem pliku Skroty.docx "
                        "(domyślnego ani własnego użytkownika)."}
            cz = (czcionki or {}).get("SKROTY") if isinstance(czcionki, dict) else None
            mg = szablony._marginesy(_marginesy_z_slownika(margins), "SKROTY")
            try:
                html = szablony.skroty_html_dopasowany(
                    sk, czcionki=cz, marginesy=mg)
            except Exception:
                return {"ok": False, "error": traceback.format_exc(limit=1)}
            return {"ok": True, "html": html, "poziom": False,
                    "zrodlo": "Skroty.docx"}
        if typ not in szablony.RENDERERY:
            return {"ok": False, "error": f"Nieznany typ raportu: {typ}"}
        txt, zrodlo = self._find_preview_txt(typ)
        if not txt:
            return {"ok": False, "error":
                    f"Nie znalazłem pliku {typ}.TXT.\n"
                    "Wskaż w kreatorze folder (źródłowy albo docelowy), "
                    "w którym leżą pliki TXT — albo uruchom najpierw proces, "
                    "a podgląd pokaże Twój dokument."}
        obiekt, stan, okres = szablony.meta_z_pliku(txt)
        mg = szablony._marginesy(_marginesy_z_slownika(margins), typ)
        cz = (czcionki or {}).get(typ) if isinstance(czcionki, dict) else None
        bez = False
        if typ in szablony.USUWA_NAZWISKA:
            try:
                bez = bool((getattr(self, "remove_names_var", None) or FakeVar(False)).get()
                           or (getattr(self, "ns_bez_nazwisk_var", None) or FakeVar(False)).get())
            except Exception:
                bez = False
        try:
            if typ == "WSKAZ1":
                html = szablony.RENDERERY[typ](txt, obiekt, stan, okres=okres,
                                               bez_nazwisk=bez, marginesy=mg,
                                               czcionki=cz)
            elif typ == "WSK_ZB":
                html = szablony.RENDERERY[typ](txt, obiekt, okres or stan,
                                              bez_nazwisk=bez, marginesy=mg,
                                              czcionki=cz)
            else:
                html = szablony.RENDERERY[typ](txt, obiekt, stan,
                                               bez_nazwisk=bez, marginesy=mg,
                                               czcionki=cz)
        except Exception:
            return {"ok": False, "error": traceback.format_exc(limit=1)}
        return {"ok": True, "html": html, "poziom": typ in szablony.POZIOMO,
                "zrodlo": zrodlo or ""}

    def _pv_win_zamkniete(self, *args, **kwargs):
        """Zamknięto okno podglądu — zapomnij referencję (patrz open_preview_window)."""
        self._pv_win = None

    def close_preview_window(self):
        """Zamyka osobne okno podglądu (wywoływane przy zamykaniu kreatora)."""
        win = getattr(self, "_pv_win", None)
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass
            self._pv_win = None
        return {"ok": True}

    def open_preview_window(self, mode="ALL"):
        """Podgląd marginesów/czcionek w OSOBNYM oknie systemowym.

        Okno można przesunąć w dowolne miejsce ekranu (poza program też)
        i dowolnie skalować — natywny pasek tytułowy systemu Windows.
        Zawartość sama się odświeża (okno pyta co ~1 s o aktualne
        marginesy/czcionki z zapisanych ustawień). Gdy wywołane ponownie,
        odświeża istniejące okno zamiast otwierać kolejne.
        """
        try:
            import webview
        except Exception as e:
            return {"ok": False, "error": f"Brak modułu pywebview: {e}"}
        self._pv_mode = "NS" if str(mode).upper() == "NS" else "ALL"
        try:
            win = getattr(self, "_pv_win", None)
            if win is not None:
                try:
                    win.show()   # daj okno na wierzch (mogło zejść za program)
                    win.evaluate_js("odswiezNatychmiast && odswiezNatychmiast()")
                    return {"ok": True, "istniejace": True}
                except Exception:
                    # okno zostało zamknięte krzyżykiem (zdarzenie closed mogło
                    # nie zdążyć wyczyścić referencji) — poniżej otwieramy nowe
                    self._pv_win = None
            self._pv_win = webview.create_window(
                "Podgląd marginesów i czcionek", html=_PV_OKNO_HTML,
                width=880, height=980, background_color="#23262d",
                js_api=self)
            try:
                # zamknięcie okna czyści referencję — żeby ponowne kliknięcie
                # przycisku otworzyło świeże okno zamiast "odświeżać" zamknięte
                self._pv_win.events.closed += self._pv_win_zamkniete
            except Exception:
                pass
            return {"ok": True}
        except Exception as e:
            self._pv_win = None
            return {"ok": False, "error": str(e)}

    def preview_window_state(self, typ):
        """Stan podglądu dla osobnego okna: raport wygenerowany na AKTUALNYCH
        marginesach i czcionkach zapisanych w ustawieniach (kreator zapisuje
        je na bieżąco, więc okno reaguje na zmiany bez restartu)."""
        from app.config import load_margins
        mode = getattr(self, "_pv_mode", "ALL")
        m = (load_margins() or {}).get(mode) or {}
        # czcionki trzymane są pod płaskim kluczem "web.czcionki.ALL"/".NS"
        # (zapisuje je kreator/zakładka Nowe Szablony przez set_values)
        cz = self.get_setting(f"web.czcionki.{mode}", None)
        if not cz and mode == "NS":
            cz = self.get_setting("web.czcionki.ALL", None)
        r = self.get_margins_preview(typ, m, cz)
        if r.get("ok"):
            try:
                from app.core import szablony
                # _marginesy zwraca (góra, prawo, dół, lewo) w cm
                mg = szablony._marginesy(_marginesy_z_slownika(m), str(typ).upper())
                r["marginesy"] = {"T": mg[0], "R": mg[1], "B": mg[2], "L": mg[3]}
            except Exception:
                pass
        return r

    def save_mapa_drop(self, nazwa, b64):
        """Mapa przeciągnięta w kreatorze (drag&drop) → plik na dysku.

        Przeglądarka nie przekazuje pełnej ścieżki upuszczonego pliku,
        więc dostajemy zawartość (base64) i zapisujemy ją sami pod
        tymczasową ścieżką; tę ścieżkę kreator wpisuje w pole 'Plik mapy'.
        """
        import base64
        import re
        try:
            nazwa = str(nazwa or "mapa").strip()
            b64 = str(b64 or "")
            if "," in b64[:64]:          # format data:image/...;base64,....
                b64 = b64.split(",", 2)[1]
            if not re.match(r"^[\w .-]+$", nazwa, re.UNICODE):
                return {"ok": False, "error": "Niedozwolona nazwa pliku."}
            if not nazwa.lower().endswith((".jpg", ".jpeg", ".png", ".tif", ".tiff")):
                return {"ok": False,
                        "error": "Mapa musi być plikiem jpg, png albo tiff."}
            folder = Path(tempfile.gettempdir()) / "forestly_mapy"
            folder.mkdir(parents=True, exist_ok=True)
            cel = folder / nazwa
            try:
                dane = base64.b64decode(b64, validate=True)
            except Exception:
                return {"ok": False,
                        "error": "Plik nie przetrwał przeciągania (uszkodzone "
                                 "dane). Użyj przycisku 'Wybierz' obok pola "
                                 "ścieżki."}
            cel.write_bytes(dane)
            # szybki test spójności — duże pliki potrafią się uciąć w moście
            try:
                from PIL import Image
                with Image.open(cel) as im:
                    im.verify()
            except Exception:
                try:
                    cel.unlink()
                except OSError:
                    pass
                return {"ok": False,
                        "error": "Plik wygląda na uszkodzony po przeciągnięciu "
                                 "(mógł się nie zmieścić w całości). Wskaż go "
                                 "przyciskiem 'Wybierz' — wtedy czytam go "
                                 "bezpośrednio z dysku."}
            return {"ok": True, "path": str(cel), "folder": str(folder)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

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
        # gdy ten folder nie ma jeszcze własnego układu — pokaż globalny
        _store = load_order_store(config_folder)
        if not (isinstance(_store.get(mode), list) and _store.get(mode)):
            _glob = (self.load_settings() or {}).get(f"pdf_order.{mode}")
            if isinstance(_glob, dict) and isinstance(_glob.get("order"), list) \
                    and _glob["order"]:
                order = _glob["order"]
                excluded = _glob.get("excluded") or []
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
        # zapas globalny — używany, gdy wyniki trafią do innego folderu
        # (kolejność i wykluczenia są wtedy stosowane też tam)
        try:
            self.set_setting(f"pdf_order.{mode}",
                             {"order": list(order or []),
                              "excluded": list(excluded or [])})
        except Exception:
            pass
        # w programie są dwa przyciski "Skonfiguruj układ PDF" (zakładka
        # 1-Click — tryb ALL — oraz zakładka Scalanie PDF — tryb PDF);
        # użytkownik konfiguruje JEDEN układ dla całego programu, więc
        # zapis trafia do obu trybów — inaczej układ ustawiony "w PDF"
        # nie działał w Pełnym Automacie (i odwrotnie)
        try:
            siostrzany = "PDF" if mode == "ALL" else "ALL"
            set_saved_template_order(config_folder, siostrzany, list(order or []))
            set_saved_excluded_templates(config_folder, siostrzany, list(excluded or []))
            self.set_setting(f"pdf_order.{siostrzany}",
                             {"order": list(order or []),
                              "excluded": list(excluded or [])})
        except Exception:
            pass
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
            # Word startowany przez COM żyje POZA drzewem procesów workera
            # i przeżywa taskkill /T — ubijamy go z rejestru, żeby nie
            # trzymał blokad na folderach wynikowych po przerwaniu zadania
            try:
                from app.core import office_guard
                office_guard.kill_registered(log=self.log)
            except Exception:
                pass
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
