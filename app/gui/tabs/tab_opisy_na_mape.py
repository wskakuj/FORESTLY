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

        tryb = _txt("mapa_zrodlo_var", "Baza MIETEK")
        if "TAKSATOR" in tryb.upper():
            tryb_kod = "taksator"
        elif "GO" in tryb.upper():
            tryb_kod = "excel"
        else:
            tryb_kod = "mietek"
        return {
            "tryb": tryb_kod,
            "excel": _txt("mapa_excel_entry"),
            "mietki": _txt("mapa_mietki_entry"),
            "mdb": _txt("mapa_mdb_entry"),
            "mapy": _txt("mapa_src_entry"),
            "kol_nr": _txt("mapa_kol_nr_entry", "N"),
            "font_mm": "2.5",          # stałe — pole usunięte
            "skala": "5000",           # stałe — pole usunięte
            "p3": "-0.25",             # stałe — zgodne z programem ukladania (bylo +0.25)
            "tylko_srodek": False,     # zgodne z programem ukladania (bylo True)
            "klucz": "TX",          # pole „Uwagi” mapy — stałe, nieedytowalne
            "a2": True,              # oba pola opisu zawsze wpisywane
            "a5": True,
        }

    # ============================================= układanie opisów na mapie
    def start_uloz_opisy(self):
        """Rozsuwa podpisy na mapach tak, aby się nie nakładały."""
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).exists():
            self.log("[UKLAD] Wskaż folder z mapami (.MAP) albo plik mapy.")
            self.update_status("Brak map", "#D83B01", animate=False)
            return
        if self.running:
            return
        self._disable_ui_for_process()
        self.log("[UKLAD] Rozsuwam podpisy na mapach (wysokość pisma %s mm, "
                 "skala 1:%s)..." % (u.get("font_mm"), u.get("skala")))
        self.set_progress(0)
        threading.Thread(target=self.run_uloz_opisy_thread, args=(u,),
                         daemon=True).start()

    # ================================================ NOWA KARTA: Układanie opisów
    def start_ukladanie_opisow(self):
        """Układanie opisów — 1:1 tak samo, jak w osobnym programie GEO-MAP_uklad.

        Stałe ustawienia (te same, co w programie, który układa dobrze):
        pismo 2,5 mm, skala 1:5000, obrót -0,25 (P3 = -0.25 grad).
        """
        src = ""
        e = getattr(self, "ukl_src_entry", None)
        if e is not None:
            try:
                src = (e.get() or "").strip()
            except Exception:
                src = ""
        if not src or not Path(src).exists():
            self.log("[UKLAD] Wskaż folder z mapami (.MAP) albo plik mapy.")
            self.update_status("Brak map", "#D83B01", animate=False)
            return
        if self.running:
            return
        self._disable_ui_for_process()
        self.log("[UKLAD] Rozsuwam opisy (pismo 2,5 mm, skala 1:5000, "
                 "obrót -0,25) — tak samo jak w programie GEO-MAP_uklad...")
        self.set_progress(0)
        threading.Thread(target=self._ukladanie_watek, args=(src,),
                         daemon=True).start()

    def _ukladanie_watek(self, src):
        import math as _m
        from pathlib import Path as _P
        try:
            from app.core import uklad_opisow as uk
            from app.core import opisy_na_mape as onm

            obrot = -0.25 * _m.pi / 200.0

            p = _P(src)
            if p.is_file() and p.suffix.lower() == ".map":
                mapy = {p.stem: p}
            elif p.is_dir():
                mapy = {}
                for f in sorted(p.rglob("*")):
                    if f.is_file() and f.suffix.lower() == ".map":
                        mapy.setdefault(f.stem, f)
            else:
                mapy = {}
            if not mapy:
                self.log("[UKLAD] Nie znaleziono plików .MAP.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return

            total = len(mapy)
            self.start_progress_tracking(total, "Układanie opisów")
            gotowe = 0
            for idx, (klucz, sciezka) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka.name
                try:
                    mapa = onm.wczytaj_mape(sciezka)
                    el = uk.uloz(mapa, wysokosc_mm=2.5, skala=5000,
                                 obrot=obrot, tylko_srodek=False)
                    if not el:
                        self.log("  • %s: brak opisów do ułożenia." % sciezka.name)
                        self.set_progress(idx / total, current_file=sciezka.name,
                                          current=idx)
                        continue
                    lines, zmiany = uk.ustaw_offsety(mapa, el, obrot_rad=obrot)
                    out = sciezka.parent / (sciezka.stem + "_ulozone.MAP")
                    out.write_bytes(onm.przelicz_naglowek(lines))
                    wew = sum(1 for e in el if e.get("wewnatrz"))
                    self.log("  • %s: opisów %d, przesunięto %d, w środku %d → %s"
                             % (sciezka.name, len(el), zmiany, wew, out.name))
                    gotowe += 1
                except InterruptedError:
                    raise
                except Exception as ex:
                    self.log("  ⚠️ %s: %s" % (sciezka.name, ex))
                self.set_progress(idx / total, current_file=sciezka.name,
                                  current=idx)
            self.log("[UKLAD] Gotowe. Map: %d, ułożonych: %d." % (total, gotowe))
            self.update_status("Ułożone opisy gotowe.", "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[UKLAD] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as ex:
            self.log("\n[UKLAD] Błąd: %s" % ex)
            traceback.print_exc()
            self.update_status("Błąd układania", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    def run_uloz_opisy_thread(self, u):
        try:
            from app.core import uklad_opisow as uk
            from app.core import opisy_na_mape as onm

            def _liczba(txt, dom):
                try:
                    return float(str(txt).replace(",", "."))
                except (TypeError, ValueError):
                    return dom

            font_mm = _liczba(u.get("font_mm"), 2.5)
            skala = _liczba(u.get("skala"), 5000)
            import math as _m
            # P3 w GEO-MAP podajemy w gradach; w pliku kąt jest w radianach
            p3 = _liczba(u.get("p3"), -0.25)
            obrot = p3 * _m.pi / 200.0
            if font_mm <= 0:
                font_mm = 2.5
            if skala <= 0:
                skala = 5000

            mapy = self._mapy_ze_sciezki(u["mapy"])
            if not mapy:
                self.log("[UKLAD] Nie znaleziono plików .MAP.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return
            total = len(mapy)
            self.start_progress_tracking(total, "Układanie opisów")
            wyniki = []
            for idx, (klucz, sciezka) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka.name
                try:
                    mapa = onm.wczytaj_mape(sciezka)
                    el = uk.uloz(mapa, wysokosc_mm=font_mm, skala=skala,
                                 obrot=obrot,
                                 tylko_srodek=u.get("tylko_srodek", False))
                    if not el:
                        wyniki.append({"mapa": sciezka.name, "opisow": 0,
                                       "zmiany": 0, "wewnatrz": 0, "plik": ""})
                    else:
                        lines, zmiany = uk.ustaw_offsety(mapa, el,
                                                         obrot_rad=obrot)
                        out = sciezka.parent
                        nazwa = (Path(sciezka.name).stem + "_ulozone.MAP")
                        (out / nazwa).write_bytes(onm.przelicz_naglowek(lines))
                        kol = {}
                        try:
                            kol = uk.policz_kolizje(mapa, lines=lines,
                                                    wysokosc_mm=font_mm,
                                                    skala=skala, obrot=obrot)
                        except Exception:
                            kol = {}
                        # PRZYCZYNY wysięgników — po co opis wyszedł na zewnątrz
                        rap = []
                        try:
                            rap = uk.raport_zewnatrz(mapa, el,
                                                     wysokosc_mm=font_mm,
                                                     skala=skala, obrot=obrot)
                            if rap:
                                import csv as _csv
                                zrodlo = Path(sciezka)
                                cel_csv = zrodlo.parent / (
                                    "Opisy na mapę - przyczyny wysięgników.csv")
                                nowy = not cel_csv.exists()
                                with open(cel_csv, "a", newline="", encoding="utf-8-sig") as fh:
                                    wr = _csv.writer(fh, delimiter=";")
                                    if nowy:
                                        wr.writerow(["mapa", "wydzielenie", "litera",
                                                     "opis", "przyczyna", "wyjaśnienie"])
                                    for r in rap:
                                        wr.writerow([zrodlo.name, r["a6"], r["litera"],
                                                     r["tekst"], r["powod"],
                                                     r["powod_tekst"]])
                        except Exception:
                            rap = []
                        wyniki.append({
                            "mapa": sciezka.name, "opisow": len(el),
                            "zmiany": zmiany,
                            "wewnatrz": sum(1 for e in el if e["wewnatrz"]),
                            "kol": kol, "rap": rap, "plik": nazwa})
                        self.log("  • %s: opisów %d, przesunięto %d, wewnątrz %d "
                                 "→ %s" % (sciezka.name, len(el), zmiany,
                                           sum(1 for e in el if e["wewnatrz"]), nazwa))
                        if rap:
                            from collections import Counter as _Cnt
                            _c = _Cnt(r["powod"] for r in rap)
                            self.log("      wysięgniki: %d — %s" % (
                                len(rap), ", ".join("%s %d" % (k, v)
                                                    for k, v in _c.most_common())))
                        if kol:
                            ostrz = "" if (kol["litery"] == 0 and kol["opis_opis"] == 0
                                           and kol["linie"] == 0) else "  ⚠️"
                            self.log("      kontrola: na literze %d, opis na opis %d, "
                                     "przecięcia linii %d%s"
                                     % (kol["litery"], kol["opis_opis"],
                                        kol["linie"], ostrz))
                except Exception as e:
                    self.log("  ⚠️ %s: %s" % (sciezka.name, e))
                self.set_progress(idx / total, current_file=sciezka.name, current=idx)
            self._uloz_raport(wyniki, u)
            self.update_status("Ułożono opisy.", "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[UKLAD] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log("\n[UKLAD] Błąd: %s" % e)
            traceback.print_exc()
            self.update_status("Błąd układania", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    def _uloz_raport(self, wyniki, u):
        linie = [
            "FORESTLY — UKŁADANIE OPISÓW NA MAPIE",
            "Data: %s" % _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Wysokość pisma: %s mm, skala 1:%s" % (u.get("font_mm"), u.get("skala")),
            "-" * 70,
        ]
        for w in wyniki:
            kol = w.get("kol") or {}
            linie.append("  • %-28s opisów: %4d, przesunięto: %4d, wewnątrz: %4d%s"
                         % (w["mapa"], w["opisow"], w["zmiany"], w["wewnatrz"],
                            (", wynik: %s" % w["plik"]) if w["plik"] else ""))
            if kol:
                linie.append("      kontrola jakości: opis na literze %d, "
                             "opis na opis %d, przecięcia linii %d"
                             % (kol.get("litery", 0), kol.get("opis_opis", 0),
                                kol.get("linie", 0)))
        linie += ["-" * 70,
                  "Razem map: %d, opisów: %d, przesunięto: %d"
                  % (len(wyniki), sum(w["opisow"] for w in wyniki),
                     sum(w["zmiany"] for w in wyniki))]
        out = Path(u["mapy"])
        if out.is_file():
            out = out.parent
        try:
            (out / "Opisy na mapę - układanie.txt").write_text(
                "\n".join(linie), encoding="utf-8-sig")
        except OSError as e:
            self.log("[UKLAD] Nie udało się zapisać raportu: %s" % e)
        self.log("\n[UKLAD] Map: %d, opisów: %d, przesunięto: %d"
                 % (len(wyniki), sum(w["opisow"] for w in wyniki),
                    sum(w["zmiany"] for w in wyniki)))

    @staticmethod
    def _zapisz_braki_csv(folder, wszystkie):
        """Zapisuje tabelę różnic do pliku „Opisy na mapę - braki.csv".

        Zwraca ścieżkę pliku albo None (gdy brak pozycji / błąd zapisu).
        """
        if not wszystkie:
            return None
        import csv as _csv
        out = Path(folder)
        if out.is_file():
            out = out.parent
        plik = out / "Opisy na mapę - braki.csv"
        try:
            with open(plik, "w", newline="", encoding="utf-8-sig") as f:
                wr = _csv.writer(f, delimiter=";")
                wr.writerow(["Mapa", "Wydzielenie", "Pole", "Co jest (mapa)",
                             "Co da reguła", "Uwaga"])
                for r in wszystkie:
                    wr.writerow([r.get("mapa", ""), r.get("wydz", ""),
                                 r.get("pole", ""), r.get("obecne", ""),
                                 r.get("nowe", ""), r.get("typ", "")])
            return plik
        except OSError:
            return None

    @staticmethod
    def _mapy_ze_sciezki(sciezka):
        """Folder -> wszystkie mapy; plik .MAP -> tylko ta jedna."""
        p = Path(sciezka)
        if p.is_file() and p.suffix.lower() == ".map":
            return {_klucz_nazwy(p.stem): p}
        return TabOpisyNaMapeMixin._mapy_w_folderze(p)

    # -------------------------------------------------- start
    def start_opisy_na_mape(self):
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).exists():
            self.log("[MAPY] Wskaż folder z mapami (.MAP) albo plik mapy.")
            self.update_status("Brak map", "#D83B01", animate=False)
            return
        if u["tryb"] == "taksator":
            if not self._bazy_ze_sciezki(u["mdb"]):
                self.log("[MAPY] Wskaż plik bazy taksatora (.mdb) — "
                         "albo folder z bazami .mdb (dopasuję po nazwie mapy).")
                self.update_status("Brak bazy .mdb", "#D83B01", animate=False)
                return
            if self.running:
                return
            self._disable_ui_for_process()
            self.log("[MAPY] Źródło: baza TAKSATORA (.mdb).")
            self.set_progress(0)
            threading.Thread(target=self.run_taksator_thread, args=(u,),
                             daemon=True).start()
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
        self.log("[MAPY] Źródło: %s."
                 % ("Excel z Forestly GO" if u["tryb"] == "excel" else "dane MIETKA"))
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
            self.update_status("Gotowe — opisy wpisane do map.",
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
    @staticmethod
    def _bazy_ze_sciezki(sciezka):
        """Folder -> wszystkie bazy .mdb; plik .mdb -> tylko ta jedna."""
        if not sciezka:
            return {}
        p = Path(sciezka)
        if p.is_file() and p.suffix.lower() == ".mdb":
            return {_klucz_nazwy(p.stem): p}
        if p.is_dir():
            return {_klucz_nazwy(q.stem): q for q in sorted(p.glob("*.mdb"))}
        return {}

    def _baza_dla_mapy(self, nazwa_mapy, bazy, cache):
        """Baza .mdb dopasowana po nazwie mapy (albo jedyna dostępna).

        Zwraca (sciezka_bazy, baza) albo (None, None).
        """
        if not bazy:
            return None, None
        # _dopasuj zwraca ŚCIEŻKĘ kandydata (nie klucz)
        sciezka = self._dopasuj(nazwa_mapy, bazy)
        if sciezka is None and len(bazy) == 1:
            sciezka = next(iter(bazy.values()))
        if sciezka is None:
            return None, None
        if sciezka not in cache:
            from app.core import opisy_na_mape_taksator as tk
            cache[sciezka] = tk.czytaj_baze(str(sciezka)) or {}
        return sciezka, cache[sciezka]

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
             "zmiany": 0, "plik": "", "blad": "", "niedopasowane": []}
        try:
            mapa, wiersze, blad = self._wiersze_dla_mapy(sciezka_mapy, zrodla, u)
            if blad:
                w["blad"] = blad
                return w
            w["poligonow"] = len(wiersze)
            w["dopasowano"] = sum(1 for r in wiersze if r["ok"])
            w["niedopasowane"] = [r["wydz"] for r in wiersze if not r["ok"]]
            if not w["dopasowano"]:
                return w
            # układanie ZAWSZE zaraz po wpisaniu (jeden przycisk):
            # każdy opis w środku wydzielenia i wspólny obrót P3
            obr = None
            try:
                import math as _m2
                obr = float(str(u.get("p3", "-0.25")).replace(",", ".")) * _m2.pi / 200.0
            except (TypeError, ValueError):
                obr = -0.25 * 3.141592653589793 / 200.0
            res = onm.zapisz_mape(mapa, mapa["path"].parent, a2=u["a2"],
                                  a5=u["a5"], wiersze=wiersze,
                                  obrot_srodek=obr)
            w["zmiany"] = res["zmiany"]
            w["plik"] = res["nazwa"]
            try:
                import math as _m3
                obr2 = float(str(u.get("p3", "-0.25")).replace(",", ".")) * _m3.pi / 200.0
            except (TypeError, ValueError):
                obr2 = -0.25 * 3.141592653589793 / 200.0
            try:
                n = self._wyśrodkuj_plik(
                    res["sciezka"], obr2,
                    font_mm=float(str(u.get("font_mm", "2.5")).replace(",", ".") or 2.5),
                    skala=float(str(u.get("skala", "5000")).replace(",", ".") or 5000),
                    rozsuwaj=bool(u.get("rozsuwaj", False)))
                self.log("  ✔ %s: wyśrodkowano i obrócono %d opisów (P3=%s grad)"
                         % (sciezka_mapy.name, n, u.get("p3")))
            except Exception as e:
                self.log("  ⚠️ %s: NIE udało się wyśrodkować/obrócić: %s"
                         % (sciezka_mapy.name, e))
        except Exception as e:
            w["blad"] = str(e)
        return w

    # ------------------------------- wyśrodkowanie i obrót gotowego pliku
    def _wyśrodkuj_plik(self, sciezka, obr, font_mm=2.5, skala=5000.0,
                        rozsuwaj=False):
        """Po wpisaniu: wyśrodkuj i obróć opisy w pliku wynikowym.

        Robimy to na GOTOWYM pliku jako osobny, widoczny krok — dzięki temu
        nawet gdyby coś w rdzeniu zawiodło po cichu, mapa i tak zostanie
        ułożona, a błąd zobaczymy w logu.
        """
        from app.core import uklad_opisow as uk
        from app.core import opisy_na_mape as onm
        mapa = onm.wczytaj_mape(sciezka)
        el = uk.uloz(mapa, wysokosc_mm=font_mm, skala=skala, obrot=obr,
                     tylko_srodek=not rozsuwaj)
        if not el:
            self.log("  ℹ️ %s: brak opisów do ułożenia." % sciezka.name)
            return 0
        lines, _zmiany = uk.ustaw_offsety(mapa, el, obrot_rad=obr)
        sciezka.write_bytes(onm.przelicz_naglowek(lines))
        return len(el)

    # -------------------------------------------------- raport
    def _opisy_raport(self, wyniki, u):
        znacznik = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        linie = [
            "FORESTLY — OPISY NA MAPĘ (GEO-MAP)",
            "Data: %s" % _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Źródło opisów: %s" % ("Excel z Forestly GO" if u["tryb"] == "excel"
                                   else "dane MIETKA"),
            "Tryb: ZAPIS map _z_opisami.MAP",
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
                if w.get("niedopasowane"):
                    linie.append("      niedopasowane (brak w źródle): %s"
                                 % ", ".join(w["niedopasowane"]))
        linie += [
            "-" * 70,
            "Razem map: %d, poligonów: %d, dopasowanych: %d, zmian pól: %d"
            % (len(wyniki), suma_pol, suma_dop, suma_zm),
        ]
        raport = "\n".join(linie)
        out = Path(u["mapy"])
        try:
            out.mkdir(parents=True, exist_ok=True)
            plik = out / "Opisy na mapę - raport.txt"
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

    def start_sprawdz_opisy_na_mape(self):
        """Sprawdza mapy i źródło — pokazuje braki do uzupełnienia/sprawdzenia."""
        u = self._opisy_na_mape_ustawienia()
        if not u["mapy"] or not Path(u["mapy"]).exists():
            self.log("[SPRAWDZ] Wskaż folder z mapami (.MAP) albo plik mapy.")
            self.update_status("Brak map", "#D83B01", animate=False)
            return
        if u["tryb"] == "taksator":
            if not self._bazy_ze_sciezki(u["mdb"]):
                self.log("[SPRAWDZ] Wskaż plik bazy taksatora (.mdb) — "
                         "albo folder z bazami .mdb.")
                self.update_status("Brak bazy .mdb", "#D83B01", animate=False)
                return
            if self.running:
                return
            self._disable_ui_for_process()
            self.log("[SPRAWDZ] Sprawdzam braki (baza TAKSATORA)...")
            self.set_progress(0)
            threading.Thread(target=self.run_taksator_braki_thread, args=(u,),
                             daemon=True).start()
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
            csvp = self._zapisz_braki_csv(u["mapy"], wszystkie)
            if csvp:
                self.log("[SPRAWDZ] Zapisano różnice do CSV: %s" % csvp)
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

    # ================================================= TAKSATOR (baza .mdb)
    def run_taksator_braki_thread(self, u):
        """Tabela braków dla trybu TAKSATOR (bez wpisywania do map)."""
        try:
            from app.core import opisy_na_mape_taksator as tk
            from app.core import opisy_na_mape as onm
            self.update_status("Czytanie bazy taksatora...", "#0078D7")
            bazy = self._bazy_ze_sciezki(u["mdb"])
            cache = {}
            mapy = self._mapy_ze_sciezki(u["mapy"])
            if not mapy:
                self.log("[SPRAWDZ] Nie znaleziono plików .MAP.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return
            total = len(mapy)
            self.start_progress_tracking(total, "Sprawdzanie braków (taksator)")
            wszystkie = []
            for idx, (klucz, sciezka) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka.name
                try:
                    sc_b, baza = self._baza_dla_mapy(sciezka.stem, bazy, cache)
                    if not baza:
                        self.log("  ⚠️ %s: brak bazy .mdb o tej nazwie — pomijam."
                                 % sciezka.name)
                        self.set_progress(idx / total, current_file=sciezka.name,
                                          current=idx)
                        continue
                    mapa = onm.wczytaj_mape(sciezka)
                    wiersze = tk.zbuduj_wiersze(mapa, baza)
                    wszystkie += tk.braki_do_przegladu(sciezka.name, wiersze)
                except Exception as e:
                    self.log("  ⚠️ %s: %s" % (sciezka.name, e))
                self.set_progress(idx / total, current_file=sciezka.name, current=idx)
            podsum = onm.podsumuj_braki(wszystkie)
            csvp = self._zapisz_braki_csv(u["mapy"], wszystkie)
            if csvp:
                self.log("[SPRAWDZ] Zapisano różnice do CSV: %s" % csvp)
            self._emit({"type": "braki", "rows": wszystkie[:2000],
                        "razem": len(wszystkie), "podsumowanie": podsum,
                        "map": len(mapy),
                        "akcja": {"task": "start_opisy_na_mape",
                                  "label": "Wpisz opisy do map"}})
            self.log("[SPRAWDZ] Pozycje do sprawdzenia/uzupełnienia: %d." % len(wszystkie))
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

    def run_taksator_thread(self, u):
        try:
            from app.core import opisy_na_mape_taksator as tk
            from app.core import opisy_na_mape as onm
            self.update_status("Czytanie bazy taksatora...", "#0078D7")
            bazy = self._bazy_ze_sciezki(u["mdb"])
            cache = {}
            mapy = self._mapy_ze_sciezki(u["mapy"])
            if not mapy:
                self.log("[TAKSATOR] Nie znaleziono plików .MAP.")
                self.update_status("Brak plików .MAP", "#D83B01", animate=False)
                return
            total = len(mapy)
            self.start_progress_tracking(total, "Opisy taksatora na mapę")
            wyniki = []
            for idx, (klucz, sciezka) in enumerate(sorted(mapy.items()), start=1):
                self.check_stop()
                self.progress_current_file = sciezka.name
                w = {"mapa": sciezka.name, "poligonow": 0, "dopasowano": 0,
                     "zmiany": 0, "plik": "", "blad": "",
                     "niedopasowane": [], "bez_oznaczenia": []}
                try:
                    sc_b, baza = self._baza_dla_mapy(sciezka.stem, bazy, cache)
                    if not baza:
                        self.log("  ⚠️ %s: brak bazy .mdb o tej nazwie — pomijam."
                                 % sciezka.name)
                        w["blad"] = "brak bazy .mdb o tej nazwie"
                        wyniki.append(w)
                        self.set_progress(idx / total, current_file=sciezka.name,
                                          current=idx)
                        continue
                    mapa = onm.wczytaj_mape(sciezka)
                    pom = []
                    wiersze = tk.zbuduj_wiersze(mapa, baza, pominiete=pom)
                    w["poligonow"] = len(wiersze)
                    w["dopasowano"] = sum(1 for r in wiersze if r["ok"])
                    w["niedopasowane"] = [r["wydz"] for r in wiersze if not r["ok"]]
                    w["bez_oznaczenia"] = [r["wydz"] for r in wiersze
                                           if r["ok"] and not r["noweA2"]]
                    if pom:
                        self.log("  ℹ️ %s: pominięto %d obiektów bez wydzielenia."
                                 % (sciezka.name, len(pom)))
                    if w["dopasowano"]:
                        res = tk.zapisz_mape(mapa, sciezka.parent, wiersze=wiersze)
                        w["zmiany"] = res["zmiany"]
                        w["plik"] = res["nazwa"]
                    self.log("  • %s: dopasowano %d/%d → %s"
                             % (sciezka.name, w["dopasowano"], w["poligonow"],
                                w.get("plik") or "(podgląd)"))
                except Exception as e:
                    w["blad"] = str(e)
                    self.log("  ⚠️ %s: %s" % (sciezka.name, e))
                wyniki.append(w)
                self.set_progress(idx / total, current_file=sciezka.name, current=idx)
            self._taksator_raport(wyniki, u)
            self.update_status("Gotowe — opisy taksatora wpisane do map.",
                               "#107C10", animate=False)
        except InterruptedError:
            self.log("\n[TAKSATOR] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:
            self.log("\n[TAKSATOR] Błąd: %s" % e)
            traceback.print_exc()
            self.update_status("Błąd wpisywania z taksatora", "#D83B01", animate=False)
        finally:
            self.running = False
            self.after(0, self.restore_all_buttons)

    def _taksator_raport(self, wyniki, u):
        znacznik = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        linie = [
            "FORESTLY — OPISY TAKSACYJNE Z TAKSATORA DO MAPY GEO-MAP",
            "Data: %s" % _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
            "Baza: %s" % u["mdb"],
            "Tryb: ZAPIS map _z_opisami.MAP",
            "-" * 70,
        ]
        sp = sd = sz = 0
        for w in wyniki:
            sp += w["poligonow"]; sd += w["dopasowano"]; sz += w["zmiany"]
            if w.get("blad"):
                linie.append("  ✗ %-28s BŁĄD: %s" % (w["mapa"], w["blad"]))
            else:
                linie.append("  • %-28s poligonów: %4d, dopasowano: %4d%s"
                             % (w["mapa"], w["poligonow"], w["dopasowano"],
                                (", zapisano: %s" % w["plik"]) if w["plik"] else ""))
                if w.get("niedopasowane"):
                    linie.append("      niedopasowane (brak w bazie): %s"
                                 % ", ".join(w["niedopasowane"]))
                if w.get("bez_oznaczenia"):
                    linie.append("      bez oznaczenia w bazie (do uzupełnienia): %s"
                                 % ", ".join(w["bez_oznaczenia"]))
        linie += ["-" * 70,
                  "Razem map: %d, poligonów: %d, dopasowanych: %d, zmian: %d"
                  % (len(wyniki), sp, sd, sz)]
        out = Path(u["mapy"])
        if out.is_file():
            out = out.parent
        try:
            out.mkdir(parents=True, exist_ok=True)
            plik = out / "Opisy na mapę - raport.txt"
            plik.write_text("\n".join(linie), encoding="utf-8-sig")
            self.log("[TAKSATOR] Raport: %s" % plik)
        except OSError as e:
            self.log("[TAKSATOR] Nie udało się zapisać raportu: %s" % e)
        self.log("\n[TAKSATOR] Map: %d, poligonów: %d, dopasowanych: %d, zmian: %d"
                 % (len(wyniki), sp, sd, sz))
        # „Otwórz folder wyników" ma prowadzić do miejsca, gdzie zapisano mapy
        self.last_output_dir = str(out)
        _zap = getattr(self, "_zapamietaj_folder_wynikow", None)
        if _zap is not None:
            try:
                _zap(out)
            except Exception:
                pass
