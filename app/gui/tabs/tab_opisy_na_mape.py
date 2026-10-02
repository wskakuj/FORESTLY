# -*- coding: utf-8 -*-
"""
Forestly — Mixin: TabOpisyNaMapeMixin

Zakładka MIETEK | Opisy na mapę (GEO-MAP). Wpisuje opisy w poligony mapy
GEO-MAP (plik .MAP): A2 („Oznaczenie") i A5 („Opis taks.").

Źródło opisów (do wyboru):
  * arkusz Excel z Forestly GO (eksport „Taksacja" — kolumna „opis_gotowy"),
  * dane MIETKA (pliki DBF: O*.DBF — wydzielenia, R*.DBF — elementy).

Dopasowanie:
  * Excel: numer porządkowy z mapy (domyślnie pole „TX") = kolumna arkusza
    (domyślnie „N" — nr_wydzielenia),
  * MIETEK: wydzielenie z mapy (A1/A6, np. „2d") = wydzielenie w mietku.

Oryginalna mapa NIE jest zmieniana — powstaje nowy plik „<NAZWA>_z_opisami.MAP".
"""

from pathlib import Path
import datetime as _dt
import re
import threading
import traceback
import unicodedata

from app.core import opisy_na_mape as onm
from app.core.wydruki import czytaj_dbf, znajdz_dbf


def _klucz_nazwy(s):
    """Nazwa do porównań: bez ogonków, małe litery, tylko [a-z0-9]."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("ł", "l").replace("Ł", "L")
    return re.sub(r"[^a-z0-9]+", "", s.lower())


class TabOpisyNaMapeMixin:
    """Mixin dla WebBackend — wpisywanie opisów do map GEO-MAP."""

    # -------------------------------------------------- dopasowanie źródeł
    @staticmethod
    def _mapy_w_folderze(folder):
        folder = Path(folder)
        mapa = {}
        for p in sorted(folder.rglob("*")):
            if p.is_file() and p.suffix.lower() == ".map":
                mapa.setdefault(_klucz_nazwy(p.stem), p)
        return mapa

    @staticmethod
    def _dopasuj(nazwa, kandydaci):
        """kandydaci: {klucz: ścieżka}. Zwraca najlepsze dopasowanie."""
        k = _klucz_nazwy(nazwa)
        if not k:
            return None
        if k in kandydaci:
            return kandydaci[k]
        # zawieranie (np. „CHORZEWO" vs „CHORZEWOWOL001")
        for kk, sc in kandydaci.items():
            if kk and (kk in k or k in kk):
                return sc
        return None

    @staticmethod
    def _arkusze_w_folderze(folder):
        folder = Path(folder)
        out = {}
        for p in sorted(folder.rglob("*.xls*")):
            if p.is_file() and not p.name.startswith("~$"):
                out.setdefault(_klucz_nazwy(p.stem), p)
        return out

    @staticmethod
    def _obreby_w_folderze(folder):
        """{klucz nazwy: folder z O*.DBF} dla wszystkich obrębów w drzewie."""
        folder = Path(folder)
        out = {}
        for dbf in sorted(folder.rglob("O*.DBF")) + sorted(folder.rglob("O*.dbf")):
            d = dbf.parent
            # nazwa obrębu to folder albo jego rodzic (np. „CHORZEWO\\WOL.001")
            for nm in (d.name, d.parent.name):
                if nm:
                    out.setdefault(_klucz_nazwy(nm), d)
        return out

    # -------------------------------------------------- tryb / kontrolki
    def _opisy_na_mape_ustawienia(self):
        def _txt(attr, default=""):
            e = getattr(self, attr, None)
            if e is None:
                return default
            try:
                v = e.get()
            except Exception:
                return default
            return (v or "").strip() or default

        def _bool(attr, default=False):
            v = getattr(self, attr, None)
            if v is None:
                return default
            try:
                return bool(v.get())
            except Exception:
                return default

        tryb = _txt("mapa_zrodlo_var", "Dane Forestly GO")
        return {
            "tryb": "mietek" if "MIET" in tryb.upper() else "excel",
            "excel": _txt("mapa_excel_entry"),
            "mietki": _txt("mapa_mietki_entry"),
            "mapy": _txt("mapa_src_entry"),
            "kol_nr": _txt("mapa_kol_nr_entry", "N"),
            "klucz": "TX",          # pole „Uwagi” mapy — stałe, nieedytowalne
            "a2": True,              # oba pola opisu zawsze wpisywane
            "a5": True,
            "dry": _bool("mapa_dry_var", False),
            "test_csv": _bool("mapa_test_csv_var", True),
        }

    # -------------------------------------------------- start
    def start_opisy_na_mape(self):
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).is_dir():
            self.log("[MAPY] Wskaż folder z mapami (.MAP).")
            self.update_status("Brak folderu z mapami", "#D83B01", animate=False)
            return
        if u["tryb"] == "excel" and (not u["excel"] or not Path(u["excel"]).is_dir()):
            self.log("[MAPY] Wskaż folder z arkuszami Excel (opisy z Forestly GO).")
            self.update_status("Brak folderu Excel", "#D83B01", animate=False)
            return
        if u["tryb"] == "mietek" and (not u["mietki"] or not Path(u["mietki"]).is_dir()):
            self.log("[MAPY] Wskaż folder z Mietkiem (pliki DBF w podfolderach).")
            self.update_status("Brak folderu Mietka", "#D83B01", animate=False)
            return
        if self.running:
            return

        self._disable_ui_for_process()
        self.log("[MAPY] Źródło: %s%s"
                 % ("Excel z Forestly GO" if u["tryb"] == "excel" else "dane MIETKA",
                    "   [TYLKO PODGLĄD — bez zapisu map]" if u["dry"] else ""))
        self.set_progress(0)
        threading.Thread(target=self.run_opisy_na_mape_thread, args=(u,),
                         daemon=True).start()

    # -------------------------------------------------- przebieg
    def run_opisy_na_mape_thread(self, u):
        try:
            self.update_status("Wczytywanie map i źródeł opisów...", "#0078D7")
            mapy = self._mapy_w_folderze(u["mapy"])
            if not mapy:
                self.log("[MAPY] Nie znaleziono plików .MAP w tym folderze.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return

            if u["tryb"] == "excel":
                zrodla = self._arkusze_w_folderze(u["excel"])
                self.log("[MAPY] Map: %d, arkuszy Excel: %d"
                         % (len(mapy), len(zrodla)))
            else:
                zrodla = self._obreby_w_folderze(u["mietki"])
                self.log("[MAPY] Map: %d, obrębów w Mietku: %d"
                         % (len(mapy), len(zrodla)))

            total = len(mapy)
            self.start_progress_tracking(total, "Opisy na mapę")
            wyniki = []
            for idx, (klucz, sciezka_mapy) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka_mapy.name
                w = self._opisy_dla_mapy(sciezka_mapy, zrodla, u)
                wyniki.append(w)
                if w.get("blad"):
                    self.log("  ⚠️ %s: %s" % (sciezka_mapy.name, w["blad"]))
                elif w["dopasowano"]:
                    self.log("  • %s: dopasowano %d/%d poligonów → %s"
                             % (sciezka_mapy.name, w["dopasowano"], w["poligonow"],
                                w.get("plik") or "(podgląd)"))
                else:
                    self.log("  ⚠️ %s: nie dopasowano żadnego poligonu "
                             "(sprawdź źródło/kolumnę numeru)." % sciezka_mapy.name)
                self.set_progress(idx / total, current_file=sciezka_mapy.name, current=idx)

            self._opisy_raport(wyniki, u)
            self.update_status(
                ("Podgląd gotowy — raport zapisany." if u["dry"]
                 else "Gotowe — opisy wpisane do map."),
                "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[MAPY] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log("\n[MAPY] Błąd: %s" % e)
            traceback.print_exc()
            self.update_status("Błąd wpisywania opisów", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    # -------------------------------------------------- źródło jednej mapy
    def _zrodlo_dla_mapy(self, sciezka_mapy, zrodla, u):
        """Dopasowuje i wczytuje źródło opisów. Zwraca (sc, zrodlo, blad)."""
        sc = self._dopasuj(sciezka_mapy.stem, zrodla)
        if sc is None:
            return None, None, ("nie znaleziono %s dla tej mapy"
                                % ("arkusza Excel" if u["tryb"] == "excel"
                                   else "obrębu w Mietku"))
        if u["tryb"] == "excel":
            x = onm.czytaj_xlsx(sc)
            if not x["naglowki"]:
                return sc, None, "arkusz jest pusty"
            return sc, {"naglowki": x["naglowki"], "wiersze": x["wiersze"]}, ""
        o_path = znajdz_dbf(sc, "O")
        r_path = znajdz_dbf(sc, "R")
        o_recs = czytaj_dbf(o_path) if o_path else []
        r_recs = czytaj_dbf(r_path) if r_path else []
        ob_o, rby = {}, {}
        for o in o_recs:
            ob_o[onm._czysc(o.get("ODDZIAL")) + onm._czysc(o.get("PODODDZ"))] = o
        for r in r_recs:
            k = onm._czysc(r.get("ODDZIAL")) + onm._czysc(r.get("PODODDZ"))
            rby.setdefault(k, []).append(r)
        return sc, {"obO": ob_o, "rby": rby}, ""

    # -------------------------------------------------- wiersze jednej mapy
    def _wiersze_dla_mapy(self, sciezka_mapy, zrodla, u):
        """Zwraca (mapa, wiersze, blad) — bez zapisu."""
        sc_zrodlo, zrodlo, blad = self._zrodlo_dla_mapy(sciezka_mapy, zrodla, u)
        if blad:
            return None, None, blad
        mapa = onm.wczytaj_mape(sciezka_mapy)
        pom = []
        wiersze = onm.zbuduj_wiersze(mapa, u["tryb"], zrodlo,
                                     klucz=u["klucz"], kol_nr=u["kol_nr"],
                                     pominiete=pom)
        if pom:
            self.log("  ℹ️ %s: pominięto %d obiektów 5310 bez oznaczenia "
                     "wydzielenia (A1 i A6 puste)." % (sciezka_mapy.name, len(pom)))
        return mapa, wiersze, ""

    # -------------------------------------------------- jedna mapa (zapis)
    def _opisy_dla_mapy(self, sciezka_mapy, zrodla, u):
        w = {"mapa": sciezka_mapy.name, "poligonow": 0, "dopasowano": 0,
             "zmiany": 0, "plik": "", "blad": ""}
        try:
            mapa, wiersze, blad = self._wiersze_dla_mapy(sciezka_mapy, zrodla, u)
            if blad:
                w["blad"] = blad
                return w
            w["poligonow"] = len(wiersze)
            w["dopasowano"] = sum(1 for r in wiersze if r["ok"])
            if u["dry"] or not w["dopasowano"]:
                return w
            res = onm.zapisz_mape(mapa, mapa["path"].parent, a2=u["a2"],
                                  a5=u["a5"], wiersze=wiersze)
            w["zmiany"] = res["zmiany"]
            w["plik"] = res["nazwa"]
        except Exception as e:
            w["blad"] = str(e)
        return w

    # -------------------------------------------------- raport
    def _opisy_raport(self, wyniki, u):
        znacznik = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        linie = [
            "FORESTLY — OPISY NA MAPĘ (GEO-MAP)",
            "Data: %s" % _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Źródło opisów: %s" % ("Excel z Forestly GO" if u["tryb"] == "excel"
                                   else "dane MIETKA"),
            "Tryb: %s" % ("TYLKO PODGLĄD (bez zapisu map)" if u["dry"]
                          else "ZAPIS map _z_opisami.MAP"),
            "Wpisywane pola: %s" % ", ".join(
                [p for p, on in (("A2", u["a2"]), ("A5", u["a5"])) if on] or ["—"]),
            "-" * 70,
        ]
        suma_pol = suma_dop = suma_zm = 0
        for w in wyniki:
            suma_pol += w["poligonow"]
            suma_dop += w["dopasowano"]
            suma_zm += w["zmiany"]
            if w.get("blad"):
                linie.append("  ✗ %-28s BŁĄD: %s" % (w["mapa"], w["blad"]))
            else:
                linie.append("  • %-28s poligonów: %4d, dopasowano: %4d%s"
                             % (w["mapa"], w["poligonow"], w["dopasowano"],
                                (", zapisano: %s" % w["plik"]) if w["plik"] else ""))
        linie += [
            "-" * 70,
            "Razem map: %d, poligonów: %d, dopasowanych: %d, zmian pól: %d"
            % (len(wyniki), suma_pol, suma_dop, suma_zm),
        ]
        raport = "\n".join(linie)
        out = Path(u["mapy"])
        try:
            out.mkdir(parents=True, exist_ok=True)
            plik = out / ("OPISY_NA_MAPE_%s.txt" % znacznik)
            plik.write_text(raport, encoding="utf-8-sig")
            self.log("[MAPY] Raport: %s" % plik)
        except OSError as e:
            self.log("[MAPY] Nie udało się zapisać raportu: %s" % e)
        self.log("\n[MAPY] Map: %d, poligonów: %d, dopasowanych: %d, zmian: %d"
                 % (len(wyniki), suma_pol, suma_dop, suma_zm))
        _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
        if _zap is not None:
            try:
                _zap(out)
            except Exception:
                pass

    # ================================================== TESTER reguły A2/A5
    def start_test_opisy_na_mape(self):
        """Porównuje regułę A2/A5 z opisami już wpisanymi w mapy (walidacja)."""
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).is_dir():
            self.log("[TEST] Wskaż folder z mapami (.MAP) — tymi, które JUŻ "
                     "mają wpisane opisy (wzorzec do porównania).")
            self.update_status("Brak folderu z mapami", "#D83B01", animate=False)
            return
        if u["tryb"] == "excel" and (not u["excel"] or not Path(u["excel"]).is_dir()):
            self.log("[TEST] Wskaż folder z arkuszami Excel (opisy z Forestly GO).")
            self.update_status("Brak folderu Excel", "#D83B01", animate=False)
            return
        if u["tryb"] == "mietek" and (not u["mietki"] or not Path(u["mietki"]).is_dir()):
            self.log("[TEST] Wskaż folder z Mietkiem (pliki DBF w podfolderach).")
            self.update_status("Brak folderu Mietka", "#D83B01", animate=False)
            return
        if not u["a2"] and not u["a5"]:
            self.log("[TEST] Zaznacz, co porównać: A2 („Oznaczenie”) "
                     "i/lub A5 („Opis taks.”).")
            self.update_status("Nic do porównania", "#D83B01", animate=False)
            return
        if self.running:
            return
        self._disable_ui_for_process()
        self.log("[TEST] Porównuję wynik reguły z opisami już wpisanymi "
                 "w mapy (źródło: %s)."
                 % ("Excel z Forestly GO" if u["tryb"] == "excel" else "dane MIETKA"))
        self.set_progress(0)
        threading.Thread(target=self.run_test_opisy_na_mape_thread, args=(u,),
                         daemon=True).start()

    def run_test_opisy_na_mape_thread(self, u):
        try:
            self.update_status("Porównywanie reguły z mapami...", "#0078D7")
            mapy = self._mapy_w_folderze(u["mapy"])
            if not mapy:
                self.log("[TEST] Nie znaleziono plików .MAP w tym folderze.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return
            if u["tryb"] == "excel":
                zrodla = self._arkusze_w_folderze(u["excel"])
            else:
                zrodla = self._obreby_w_folderze(u["mietki"])
            total = len(mapy)
            self.start_progress_tracking(total, "Tester opisów na mapę")
            wyniki = []
            for idx, (klucz, sciezka_mapy) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka_mapy.name
                w = self._test_dla_mapy(sciezka_mapy, zrodla, u)
                wyniki.append(w)
                if w.get("blad"):
                    self.log("  ⚠️ %s: %s" % (sciezka_mapy.name, w["blad"]))
                else:
                    czesci = []
                    for pole in ("A2", "A5"):
                        r = w["wyniki"].get(pole)
                        if r:
                            czesci.append("%s: zgodne %d, rozbieżne %d"
                                          % (pole, r["zgodne"], r["rozbiezne"]))
                    self.log("  • %s: %s"
                             % (sciezka_mapy.name, "  |  ".join(czesci) or "—"))
                self.set_progress(idx / total, current_file=sciezka_mapy.name,
                                  current=idx)
            self._test_raport(wyniki, u)
            self.update_status("Test gotowy — raport zapisany.", "#107C10",
                               animate=False)
        except InterruptedError:
            self.log("\n[TEST] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log("\n[TEST] Błąd: %s" % e)
            traceback.print_exc()
            self.update_status("Błąd testu reguły", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    def _test_dla_mapy(self, sciezka_mapy, zrodla, u):
        w = {"mapa": sciezka_mapy.name, "poligonow": 0, "dopasowano": 0,
             "wyniki": {}, "blad": ""}
        try:
            mapa, wiersze, blad = self._wiersze_dla_mapy(sciezka_mapy, zrodla, u)
            if blad:
                w["blad"] = blad
                return w
            w["poligonow"] = len(wiersze)
            w["dopasowano"] = sum(1 for r in wiersze if r["ok"])
            for pole, on in (("A2", u["a2"]), ("A5", u["a5"])):
                if on:
                    w["wyniki"][pole] = onm.podsumuj_test(wiersze, pole)
        except Exception as e:
            w["blad"] = str(e)
        return w

    def _test_raport(self, wyniki, u):
        znacznik = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        pola = [p for p, on in (("A2", u["a2"]), ("A5", u["a5"])) if on]
        linie = [
            "FORESTLY — TESTER REGUŁY A2/A5 (opisy na mapę GEO-MAP)",
            "Data: %s" % _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Źródło opisów: %s" % ("Excel z Forestly GO" if u["tryb"] == "excel"
                                   else "dane MIETKA"),
            "Porównywane pola: %s" % (", ".join(pola) or "—"),
            "Porównanie: reguła (nowe) vs to, co JEST w mapie (obecne).",
            "Różnice dzielone są na: rozbieżność (merytoryczna), literówka, "
            "interpunkcja, spacje, wielkość liter.",
            "=" * 74,
        ]
        csv = [["mapa", "wydzielenie", "pole", "rodzaj", "obecne", "nowe", "status"]]
        drobne_kat = ("literówka", "interpunkcja", "spacje", "wielkość liter")
        suma = {p: {"zgodne": 0, "rozbieżność": 0, "literówka": 0,
                    "interpunkcja": 0, "spacje": 0, "wielkość liter": 0,
                    "niedopasowane": 0, "puste": 0} for p in pola}
        for w in wyniki:
            if w.get("blad"):
                linie.append("  ✗ %s — BŁĄD: %s" % (w["mapa"], w["blad"]))
                continue
            linie.append("")
            linie.append("■ %s — poligonów: %d, dopasowanych: %d"
                         % (w["mapa"], w["poligonow"], w["dopasowano"]))
            for pole in pola:
                r = w["wyniki"].get(pole)
                if not r:
                    continue
                suma[pole]["zgodne"] += r["zgodne"]
                suma[pole]["niedopasowane"] += r["niedopasowane"]
                suma[pole]["puste"] += r["puste"]
                for k in drobne_kat + ("rozbieżność",):
                    suma[pole][k] += r["kategorie"].get(k, 0)
                istotne = r["kategorie"].get("rozbieżność", 0)
                drobne = sum(r["kategorie"].get(k, 0) for k in drobne_kat)
                linie.append("   %s: zgodne: %d, puste w mapie: %d, "
                             "niedopasowane: %d"
                             % (pole, r["zgodne"], r["puste"], r["niedopasowane"]))
                linie.append("      rozbieżności merytoryczne: %d, drobne "
                             "(literówki/spacje/interpunkcja/wielkość liter): %d"
                             % (istotne, drobne))
                for roz in r["rozbieznosci"]:
                    linie.append("      • [%s] %-8s obecne: %-24s | nowe: %s"
                                 % (roz["rodzaj"], roz["wydz"],
                                    roz["obecne"], roz["nowe"]))
                    csv.append([w["mapa"], roz["wydz"], pole, roz["rodzaj"],
                                roz["obecne"], roz["nowe"], roz["status"]])
        linie += ["", "=" * 74]
        for pole in pola:
            s = suma[pole]
            linie.append("RAZEM %s — zgodne: %d, puste w mapie: %d, "
                         "niedopasowane: %d"
                         % (pole, s["zgodne"], s["puste"], s["niedopasowane"]))
            linie.append("   rozbieżność: %d, literówka: %d, interpunkcja: %d, "
                         "spacje: %d, wielkość liter: %d"
                         % (s["rozbieżność"], s["literówka"], s["interpunkcja"],
                            s["spacje"], s["wielkość liter"]))
        linie.append("")
        linie.append("Rozbieżność merytoryczna ≠ od razu błąd reguły — sprawdź, "
                     "czy mapa (wzorzec) sama nie jest błędna.")
        raport = "\n".join(linie)
        out = Path(u["mapy"])
        try:
            out.mkdir(parents=True, exist_ok=True)
            plik = out / ("TEST_OPISY_NA_MAPE_%s.txt" % znacznik)
            plik.write_text(raport, encoding="utf-8-sig")
            self.log("[TEST] Raport: %s" % plik)
            if u["test_csv"] and len(csv) > 1:
                import csv as _csv
                cplik = out / ("TEST_ROZBIEZNOSCI_%s.csv" % znacznik)
                with open(cplik, "w", newline="", encoding="utf-8-sig") as fh:
                    _csv.writer(fh, delimiter=";").writerows(csv)
                self.log("[TEST] Rozbieżności (CSV): %s" % cplik)
        except OSError as e:
            self.log("[TEST] Nie udało się zapisać raportu: %s" % e)
        for pole in pola:
            s = suma[pole]
            drobne = sum(s[k] for k in drobne_kat)
            self.log("[TEST] %s — zgodne: %d, rozbieżność: %d, drobne: %d, "
                     "puste: %d, niedopasowane: %d"
                     % (pole, s["zgodne"], s["rozbieżność"], drobne,
                        s["puste"], s["niedopasowane"]))
        _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
        if _zap is not None:
            try:
                _zap(out)
            except Exception:
                pass

    # ================================================ SPRAWDZANIE braków
    def start_sprawdz_opisy_na_mape(self):
        """Sprawdza mapy i źródło — pokazuje braki do uzupełnienia/sprawdzenia."""
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).is_dir():
            self.log("[SPRAWDZ] Wskaż folder z mapami (.MAP).")
            self.update_status("Brak folderu z mapami", "#D83B01", animate=False)
            return
        if u["tryb"] == "excel" and (not u["excel"] or not Path(u["excel"]).is_dir()):
            self.log("[SPRAWDZ] Wskaż folder z arkuszami Excel.")
            self.update_status("Brak folderu Excel", "#D83B01", animate=False)
            return
        if u["tryb"] == "mietek" and (not u["mietki"] or not Path(u["mietki"]).is_dir()):
            self.log("[SPRAWDZ] Wskaż folder z Mietkiem (pliki DBF).")
            self.update_status("Brak folderu Mietka", "#D83B01", animate=False)
            return
        if self.running:
            return
        self._disable_ui_for_process()
        self.log("[SPRAWDZ] Sprawdzam mapy i źródło — szukam braków...")
        threading.Thread(target=self.run_sprawdz_opisy_na_mape_thread,
                         args=(u,), daemon=True).start()

    def run_sprawdz_opisy_na_mape_thread(self, u):
        try:
            self.update_status("Sprawdzanie braków...", "#0078D7")
            mapy = self._mapy_w_folderze(u["mapy"])
            if not mapy:
                self.log("[SPRAWDZ] Nie znaleziono plików .MAP w tym folderze.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return
            zrodla = (self._arkusze_w_folderze(u["excel"]) if u["tryb"] == "excel"
                      else self._obreby_w_folderze(u["mietki"]))
            total = len(mapy)
            self.start_progress_tracking(total, "Sprawdzanie braków")
            wszystkie = []
            for idx, (klucz, sciezka_mapy) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka_mapy.name
                try:
                    mapa, wiersze, blad = self._wiersze_dla_mapy(sciezka_mapy,
                                                                 zrodla, u)
                    if blad:
                        self.log("  ⚠️ %s: %s" % (sciezka_mapy.name, blad))
                    else:
                        wszystkie += onm.braki_do_przegladu(
                            sciezka_mapy.name, wiersze, u["a2"], u["a5"])
                except Exception as e:
                    self.log("  ⚠️ %s: %s" % (sciezka_mapy.name, e))
                self.set_progress(idx / total, current_file=sciezka_mapy.name,
                                  current=idx)

            podsum = onm.podsumuj_braki(wszystkie)
            self._emit({"type": "braki", "rows": wszystkie[:2000],
                        "razem": len(wszystkie),
                        "podsumowanie": podsum,
                        "map": len(mapy),
                        "akcja": {"task": "start_opisy_na_mape",
                                  "label": "Wpisz opisy do map"}})
            if wszystkie:
                self.log("[SPRAWDZ] Pozycje do sprawdzenia/uzupełnienia: %d (%s)."
                         % (len(wszystkie),
                            ", ".join("%s: %d" % (k, v)
                                      for k, v in podsum.items())))
            else:
                self.log("[SPRAWDZ] Brak uwag — wszystko dopasowane i zgodne.")
            self.update_status("Sprawdzanie gotowe.", "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[SPRAWDZ] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log("\n[SPRAWDZ] Błąd: %s" % e)
            traceback.print_exc()
            self.update_status("Błąd sprawdzania", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)
