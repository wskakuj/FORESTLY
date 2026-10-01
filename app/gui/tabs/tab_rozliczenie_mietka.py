# -*- coding: utf-8 -*-
"""
Forestly — Mixin: TabRozliczenieMietkaMixin

Zakładka MIETEK | Czyszczenie rejestru. Wskazujesz folder z mietkiem (pliki
DBF mogą leżeć w podfolderach, także głębiej), a program dla każdego obrębu:

  * usuwa DZIAŁKI bez przypisanego pododdziału (litery),
  * usuwa WŁAŚCICIELI / pozycje rejestrowe bez żadnej rozliczonej działki,
  * opcjonalnie ODŚWIEŻA już wygenerowane raporty (nowe szablony: HTML + PDF),
  * tworzy raport usuniętych pozycji (TXT).

Tryb „Tylko raport” niczego nie usuwa — służy do podglądu.
"""

from app.core.leniwe_importy import leniwy_modul
ctk = leniwy_modul("customtkinter")
from pathlib import Path
import datetime as _dt
import threading
import traceback

from app.config import load_margins
from app.core import rozliczenie_mietka as rozl


class TabRozliczenieMietkaMixin:
    """Mixin dla WebBackend — czyszczenie rejestru mietka z nierozliczonych."""

    # ------------------------------------------------ foldery raportów
    def _znajdz_folder_raportow(self, obr, root, raporty_root):
        """Folder z wygenerowanymi raportami (nowe szablony) dla obrębu."""
        obr, root = Path(obr), Path(root)
        kand = []
        if raporty_root:
            rp = Path(raporty_root)
            wies = obr.parent.name if obr.parent != root else obr.name
            kand += [rp / wies / "nowe szablony", rp / "nowe szablony", rp]
        for base in (obr, obr.parent):
            if base:
                kand.append(Path(base) / "nowe szablony")
                kand.append(Path(base))
        for k in kand:
            try:
                if k and k.is_dir() and (list(k.glob("*.html"))
                                         or (k / "pdf").is_dir()):
                    return k
            except OSError:
                continue
        return None

    def _marginesy_i_czcionki(self):
        conv = getattr(self, "_marginesy_z_slownika", None)
        margins = {}
        if conv is not None:
            for tryb in ("NS", "ALL"):
                try:
                    margins = conv(load_margins().get(tryb) or {})
                except Exception:
                    margins = {}
                if margins:
                    break
        czc = None
        get = getattr(self, "get_setting", None)
        if get is not None:
            czc = get("web.czcionki.NS", None)
            if not isinstance(czc, dict) or not czc:
                czc = get("web.czcionki.ALL", None)
        if not isinstance(czc, dict):
            czc = {}
        return margins, czc

    # ------------------------------------------------ start
    def start_czyszczenie_rejestru(self):
        root_e = getattr(self, "rozl_mietki_root_entry", None)
        root_raw = root_e.get().strip() if root_e is not None else ""
        if not root_raw or not Path(root_raw).exists():
            self.log("[CZYSZCZENIE] Wskaż istniejący folder z mietkiem "
                     "(pliki DBF — także w podfolderach).")
            self.update_status("Brak folderu", "#D83B01", animate=False)
            return
        if self.running:
            return

        out_e = getattr(self, "rozl_mietki_out_entry", None)
        out_raw = out_e.get().strip() if out_e is not None else ""
        dry_var = getattr(self, "rozl_mietki_dry_var", None)
        dry = bool(dry_var.get()) if dry_var is not None else False
        odswiez_var = getattr(self, "rozl_mietki_odswiez_var", None)
        odswiez = bool(odswiez_var.get()) if odswiez_var is not None else False

        self._disable_ui_for_process()
        self.log(f"[CZYSZCZENIE] Folder: {root_raw}\n"
                 f"Tryb: {'TYLKO RAPORT (podgląd — bez usuwania)' if dry else 'USUWANIE nierozliczonych (kopie .BAK)'}"
                 + (f"\nOdświeżanie raportów: {'TAK' if odswiez else 'nie'}"
                    + (f" ({out_raw})" if out_raw else "")))
        self.set_progress(0)
        threading.Thread(
            target=self.run_czyszczenie_rejestru_thread,
            args=(root_raw, out_raw, dry, odswiez),
            daemon=True,
        ).start()

    def run_czyszczenie_rejestru_thread(self, root_str, out_str, dry, odswiez):
        try:
            self.update_status("Szukanie obrębów i sprawdzanie rejestru...",
                               "#0078D7")
            root = Path(root_str)
            obreby = rozl.znajdz_obreby(root)
            if not obreby:
                self.log("[CZYSZCZENIE] Nie znaleziono plików DBF w tym "
                         "folderze ani w podfolderach.")
                self.update_status("Brak plików DBF", "#D83B01", animate=False)
                return

            margins, czc = self._marginesy_i_czcionki()
            total = len(obreby)
            self.start_progress_tracking(total, "Czyszczenie rejestru")
            wyniki = []
            for idx, obr in enumerate(obreby, start=1):
                self.check_stop()
                self.progress_current_file = obr.name
                try:
                    w = rozl.przeczysc_obreb(obr, usun=not dry)
                except Exception as e:
                    self.log(f"  ❌ {obr.name}: błąd — {e}")
                    traceback.print_exc()
                    w = {"obreb": obr.name, "d_rekordow": 0, "w_rekordow": 0,
                         "d_usuniete": [], "w_usuniete": [], "bak": [],
                         "blad": f"błąd odczytu/zapisu: {e}"}
                # odświeżenie raportów (tylko przy realnym usuwaniu i zmianach)
                w["raporty"] = []
                if (odswiez and not dry and not w.get("blad")
                        and (w["d_usuniete"] or w["w_usuniete"])):
                    rdir = self._znajdz_folder_raportow(obr, root, out_str)
                    if rdir:
                        try:
                            w["raporty"] = rozl.odswiez_raporty(
                                obr, rdir, margins=margins, czcionki=czc)
                            self.log(f"  ↻ {obr.name}: odświeżono raporty "
                                     f"({', '.join(w['raporty']) or '—'}) → {rdir}")
                        except Exception as e:
                            self.log(f"  ⚠️ {obr.name}: nie udało się odświeżyć "
                                     f"raportów ({e}).")
                    else:
                        self.log(f"  ℹ️ {obr.name}: nie znaleziono folderu "
                                 f"raportów (nowe szablony) — pomijam odświeżanie.")
                wyniki.append(w)
                nd, nw = len(w["d_usuniete"]), len(w["w_usuniete"])
                if w.get("blad"):
                    self.log(f"  ⚠️ {obr.name}: {w['blad']}")
                elif nd or nw:
                    self.log(f"  • {obr.name}: "
                             + ("usunięto " if not dry else "do usunięcia ")
                             + f"{nd} działek, {nw} właścicieli/pozycji"
                             + ("  (kopie .BAK)" if w["bak"] else ""))
                else:
                    self.log(f"  ✓ {obr.name}: rejestr rozliczony — nic do usunięcia.")
                self.set_progress(idx / total, current_file=obr.name, current=idx)

            raport = rozl.zbuduj_raport(wyniki, dry_run=dry)
            znacznik = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            plik = root / f"CZYSZCZENIE_REJESTRU_{znacznik}.txt"
            try:
                plik.write_text(raport, encoding="utf-8-sig")
            except OSError as e:
                self.log(f"[CZYSZCZENIE] Nie udało się zapisać raportu: {e}")
            self.last_output_dir = root
            _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
            if _zap is not None:
                _zap(root)

            suma_d = sum(len(w["d_usuniete"]) for w in wyniki)
            suma_w = sum(len(w["w_usuniete"]) for w in wyniki)
            ile_bak = sum(len(w["bak"]) for w in wyniki)
            if suma_d or suma_w:
                self.log(f"\n[CZYSZCZENIE] {'DO USUNIĘCIA' if dry else 'USUNIĘTO'}: "
                         f"{suma_d} działek, {suma_w} właścicieli/pozycji "
                         f"w {total} obrębach."
                         + (f" Kopie .BAK: {ile_bak}." if ile_bak else ""))
            else:
                self.log(f"\n[CZYSZCZENIE] Nic do usunięcia — rejestr rozliczony "
                         f"({total} obrębów).")
            self.log(f"[CZYSZCZENIE] Raport: {plik}")
            self.update_status(
                ("Podgląd gotowy — raport zapisany." if dry
                 else "Gotowe — rejestr oczyszczony, raport zapisany."),
                "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[CZYSZCZENIE] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log(f"\n[CZYSZCZENIE] Błąd: {e}")
            traceback.print_exc()
            self.update_status("Błąd czyszczenia rejestru", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)
