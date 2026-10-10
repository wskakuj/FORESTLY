# -*- coding: utf-8 -*-
"""
Forestly — Mixin: TabRozliczanieMietekMixin

Zakładka „ROZLICZANIE | Rozliczanie + MIETEK" — zbudowana na wzór
„MIETEK | Pełny Automat (1-Click)": wszystkie ustawienia w jednej zakładce,
kontrolka „dashboard" pokazująca kroki, jeden przycisk „Rozpocznij proces".

Program przechodzi po kolei:

    1. Generowanie VAL z map (.MAP),
    2. Rozliczanie powierzchni (ewidencja XLS + pliki VAL),
    3. Tworzenie mietków i wpisanie krzyżówek,
    4. Zestawienie zbiorcze (opcjonalnie — przełącznik w zakładce).

Wynik trafia do jednego nowego folderu ROZLICZANIE_MIETEK_<data_godzina>
(podfoldery MAPY, VAL, EWIDENCJA, ROZLICZONE, MIETKI).

Kod działa w OBU wersjach programu (klasycznej i webowej) — kontrolki czytamy
obronnie, a postęp kroków idzie przez update_dashboard()/reset_dashboard(),
które są w obu wersjach.
"""

from pathlib import Path
import datetime as _dt
import re
import shutil
import threading
import traceback


def _mb():
    """Moduł okien dialogowych (tkinter.messagebox).

    W wersji webowej tkinter.messagebox jest podmieniany na dialogi
    przeglądarki, więc zwykłe jego użycie działa w obu wersjach programu.
    """
    try:
        import tkinter.messagebox as _m
        return _m
    except Exception:                                       # noqa: BLE001
        class _Atrapa:
            @staticmethod
            def askyesno(tytul, tresc, **k):
                return True

            @staticmethod
            def showinfo(tytul, tresc, **k):
                pass

            @staticmethod
            def showerror(tytul, tresc, **k):
                pass

            @staticmethod
            def showwarning(tytul, tresc, **k):
                pass

        return _Atrapa


def _rm_resource_path(nazwa):
    """Ścieżka do wbudowanego zasobu programu (np. folderu „pusty")."""
    try:
        from app.core.word_worker import get_resource_path
        return get_resource_path(nazwa)
    except Exception:                                       # noqa: BLE001
        import sys as _sys
        kandydaci = []
        if getattr(_sys, "frozen", False):
            kandydaci.append(Path(_sys.executable).resolve().parent / nazwa)
        meipass = getattr(_sys, "_MEIPASS", None)
        if meipass:
            kandydaci.append(Path(meipass) / nazwa)
        kandydaci.append(Path(__file__).resolve().parent.parent.parent.parent / nazwa)
        kandydaci.append(Path.cwd() / nazwa)
        for k in kandydaci:
            if k.exists():
                return k
        return kandydaci[0]


def _rm_pliki(wpis, rozszerzenia):
    """Wpis użytkownika -> lista plików o danych rozszerzeniach.

    Obsługuje: jeden plik, kilka plików rozdzielonych średnikiem (tak zwraca
    wybór wielokrotny) oraz folder (bierze wszystkie pasujące pliki, także
    z podfolderów).
    """
    rozszerzenia = {r.lower() for r in rozszerzenia}
    czesci = [c.strip().strip('"') for c in re.split(r"[;\n]+", str(wpis or ""))
              if c.strip()]
    pliki, widziane = [], set()

    def _dodaj(p):
        try:
            klucz = str(p.resolve()).lower()
        except OSError:
            klucz = str(p).lower()
        if klucz not in widziane:
            widziane.add(klucz)
            pliki.append(p)

    for c in czesci:
        p = Path(c)
        if p.is_dir():
            for q in sorted(p.rglob("*")):
                if q.is_file() and q.suffix.lower() in rozszerzenia:
                    _dodaj(q)
        elif p.is_file() and p.suffix.lower() in rozszerzenia:
            _dodaj(p)
    return pliki


class TabRozliczanieMietekMixin:
    """Mixin dla WebBackend oraz ModernApp."""

    # ------------------------------------------------------------- ustawienia
    def _rm_ustawienia(self):
        def _txt(attr, default=""):
            e = getattr(self, attr, None)
            if e is None:
                return default
            try:
                v = e.get()
            except Exception:                               # noqa: BLE001
                return default
            return (v or "").strip() or default

        def _bool(attr, default=False):
            v = getattr(self, attr, None)
            if v is None:
                return default
            try:
                return bool(v.get())
            except Exception:                               # noqa: BLE001
                return default

        return {
            "mapy": _txt("rm_mapy_entry"),
            "ewid": _txt("rm_ewid_entry"),
            "out": _txt("rm_out_entry"),
            "wojew": _txt("rm_wojew_entry", "10"),
            "powiat": _txt("rm_powiat_entry"),
            "stan": _txt("rm_stan_entry", "01.01.2023"),
            "obod": _txt("rm_obod_entry", "01.01.2023"),
            "obdo": _txt("rm_obdo_entry", "31.12.2032"),
            "nrws": _txt("rm_nrws_entry", "1"),
            "rokz": _txt("rm_rokz_entry", "19"),
            "zestaw": _bool("rm_zestaw_var", True),
        }

    def _rm_ustaw_stan(self, zakonczone=False, blad=False, komunikat=""):
        """Zapisuje stan procesu — okno czyta go wprost (rm_diag)."""
        self._rm_stan = {"zakonczone": bool(zakonczone), "blad": bool(blad),
                         "komunikat": str(komunikat)}

    # ------------------------------------------------------- diagnostyka
    def rm_diag(self):
        """Co widzi PROGRAM (backend) — do pokazania w oknie kreatora.

        Odpowiada na pytania: czy ta wersja programu zna zadanie, czy nie
        trwa inne zadanie, jakie ścieżki dotarły z okna i co ostatnio
        zapisano w dzienniku.
        """
        try:
            from app.config import CURRENT_VERSION as _w
        except Exception:                                   # noqa: BLE001
            _w = "?"
        try:
            zadania = list(self._task_map().keys())
            ma = "start_rozliczanie_mietki" in zadania
        except Exception as e:                              # noqa: BLE001
            ma = False
            zadania = ["błąd listy zadań: %s" % e]
        try:
            u = self._rm_ustawienia()
        except Exception as e:                              # noqa: BLE001
            u = {"mapy": "?%s" % e, "ewid": "", "out": ""}
        buf = getattr(self, "_log_buf", None) or []
        import time as _t
        _lp = getattr(self, "_last_poll_ts", 0) or 0
        return {
            "ok": True,
            "wersja": _w,
            "odpytanie_s": (round(_t.time() - _lp, 1) if _lp else -1),
            "zadanie_jest": bool(ma),
            "trwa": bool(getattr(self, "running", False)),
            "pola": {"mapy": u.get("mapy", ""), "ewid": u.get("ewid", ""),
                     "out": u.get("out", "")},
            "logi": [str(x) for x in buf[-40:]],
            "stan": getattr(self, "_rm_stan", None) or {},
        }

    # ------------------------------------------------------------------ start
    def start_rozliczanie_mietki(self):
        """Uruchamia cały proces: VAL → rozliczanie → mietki → zestawienie."""
        u = self._rm_ustawienia()
        self._rm_ustaw_stan()
        brak = []
        if not u["mapy"]:
            brak.append("mapy GEO-MAP (.MAP)")
        if not u["ewid"]:
            brak.append("ewidencja (pliki XLS/XLSX)")
        if not u["out"]:
            brak.append("folder wynikowy")
        if brak:
            self.log("[R+M] Nie uruchamiam — brakuje: %s." % ", ".join(brak))
            self.update_status("Brakuje: %s" % ", ".join(brak),
                               "#D83B01", animate=False)
            self._rm_ustaw_stan(zakonczone=True, blad=True,
                                komunikat="Brakuje: %s" % ", ".join(brak))
            return
        if self.running:
            # NIE wychodzimy po cichu — inaczej użytkownik widzi tylko ogólny
            # „błąd", a w dzienniku nie ma ani jednego wpisu.
            self.log("[R+M] Nie uruchamiam — inne zadanie jest już w toku. "
                     "Kliknij „Przerwij zadanie” (na ekranie postępu) albo "
                     "poczekaj, aż się skończy, i spróbuj ponownie.")
            self.update_status("Inne zadanie jest już w toku", "#D83B01",
                               animate=False)
            self._rm_ustaw_stan(zakonczone=True, blad=True,
                                komunikat="Inne zadanie jest już w toku")
            return
        self._disable_ui_for_process()
        try:
            self.reset_dashboard()
        except Exception:                                   # noqa: BLE001
            pass
        self.set_progress(0)
        self.log("[R+M] Start procesu: VAL z map → rozliczanie → mietki"
                 + (" → zestawienie zbiorcze" if u.get("zestaw") else "")
                 + ".")
        self.log("[R+M] Ustawienia: mapy=%s | ewidencja=%s | wynik=%s"
                 % (u["mapy"], u["ewid"], u["out"]))
        threading.Thread(target=self.run_rozliczanie_mietki_thread, args=(u,),
                         daemon=True).start()

    # ------------------------------------------------------- trzymanie stanu
    def _rm_trzymaj(self):
        """Po każdym pod-kroku przywracamy stan „trwa zadanie".

        Pod-kroki (run_rozliczanie_thread, run_tworzenie_mietkow_thread...)
        same kończą się ustawieniem self.running = False i odblokowaniem
        przycisków. W procesie chcemy, żeby blokada trwała do końca.
        """
        self.running = True
        self.stop_event.clear()
        try:
            self._emit({"type": "state", "running": True})
        except Exception:                                   # noqa: BLE001
            pass

    def _rm_krok(self, index, status, text=None):
        """Bezpieczne ustawienie stanu kroku w dashboardzie."""
        try:
            self.update_dashboard(index, status, text)
        except Exception:                                   # noqa: BLE001
            pass

    # ------------------------------------------------------------------- praca
    def run_rozliczanie_mietki_thread(self, u):
        _orig_restore = None
        try:
            messagebox = _mb()

            from app.core import generowanie_val as gv
            from app.core import literacja as lit
            try:
                from app.gui.tabs.tab_rozliczanie import _pary_obrebow
            except Exception:                               # noqa: BLE001
                _pary_obrebow = None

            self.update_status("Rozliczanie + MIETEK...", "#0078D7")

            # WAŻNE: pod-kroki (rozliczanie, mietki, zestawienie) kończą się
            # ustawieniem self.running = False i wysłaniem do interfejsu
            # „zadanie zakończone". W naszym procesie to nieprawda — trwa
            # dopiero krok drugi z czterech. Podmieniamy więc na czas procesu
            # tę funkcję na pustą, żeby okno nie ogłaszało końca w połowie.
            _orig_restore = self.restore_all_buttons
            self.restore_all_buttons = lambda *a, **k: None

            # ---------------- foldery robocze -------------------------------
            baza = Path(u["out"])
            baza.mkdir(parents=True, exist_ok=True)
            stempel = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            praca = baza / ("ROZLICZANIE_MIETEK_" + stempel)
            _n = 2
            while praca.exists():
                praca = baza / ("ROZLICZANIE_MIETEK_%s_%d" % (stempel, _n))
                _n += 1
            mapy_dir = praca / "MAPY"
            val_dir = praca / "VAL"
            ewid_dir = praca / "EWIDENCJA"
            rozl_dir = praca / "ROZLICZONE"
            mietki_dir = praca / "MIETKI"
            for d in (mapy_dir, val_dir, ewid_dir, rozl_dir, mietki_dir):
                d.mkdir(parents=True, exist_ok=True)
            self.last_output_dir = praca
            try:
                self._zapamietaj_folder_wynikow(praca)
            except Exception:                               # noqa: BLE001
                pass
            self.log("[R+M] Nowy folder roboczy: %s" % praca)

            # ---------------- zebranie materiału ----------------------------
            mapy, ostrz = lit.zbierz_pliki_wejsciowe(
                u["mapy"], katalog_roboczy=praca / "_ZIP",
                log=lambda m: self.log("[R+M] %s" % m))
            for o in ostrz:
                self.log("[R+M] %s" % o)
            if not mapy:
                self._rm_krok(0, "error", "brak map")
                messagebox.showerror("Rozliczanie + MIETEK",
                                     "Nie znalazłem żadnej mapy .MAP.")
                return
            for m in mapy:
                try:
                    shutil.copy2(m, mapy_dir / m.name)
                except OSError as e:
                    self.log("[R+M] Nie skopiowano %s: %s" % (m.name, e))
            self.log("[R+M] Map do przetworzenia: %d" % len(mapy))

            ewid = _rm_pliki(u["ewid"], {".xls", ".xlsx"})
            if not ewid:
                self._rm_krok(0, "error", "brak ewidencji")
                messagebox.showerror("Rozliczanie + MIETEK",
                                     "Nie znalazłem plików ewidencji (.xls/.xlsx).")
                return
            for e in ewid:
                try:
                    shutil.copy2(e, ewid_dir / e.name)
                except OSError as exc:
                    self.log("[R+M] Nie skopiowano %s: %s" % (e.name, exc))
            self.log("[R+M] Plików ewidencji: %d" % len(ewid))

            # ================= KROK 1: VAL ==================================
            self._rm_krok(0, "running", "Generowanie VAL...")
            self.update_status("Krok 1/4 — generowanie VAL...", "#0078D7")
            self.log("[R+M] KROK 1/4 — generowanie plików VAL z map.")
            ok_val = 0
            for m in sorted(p for p in mapy_dir.iterdir() if p.is_file()):
                if self.stop_event.is_set():
                    self._rm_krok(0, "error", "przerwano")
                    self.log("[R+M] Przerwano w kroku 1.")
                    return
                try:
                    r = gv.generuj_val(m, log=lambda s: self.log("[R+M][VAL] %s" % s))
                    if isinstance(r, dict) and r.get("ok"):
                        ok_val += 1
                except Exception as e:                          # noqa: BLE001
                    self.log("[R+M][VAL] Błąd przy %s: %s" % (m.name, e))
            for v in list(mapy_dir.glob("*.VAL")) + list(mapy_dir.glob("*.val")):
                try:
                    shutil.move(str(v), str(val_dir / v.name))
                except OSError:
                    pass
            self.log("[R+M] Krok 1 gotowy: plików .VAL = %d" % ok_val)
            if ok_val == 0:
                self._rm_krok(0, "error", "0 plików VAL")
                messagebox.showerror(
                    "Rozliczanie + MIETEK",
                    "Nie udało się utworzyć żadnego pliku .VAL.\n\n"
                    "Sprawdź, czy mapy mają działki (pole A2) i wydzielenia "
                    "(pole A1) oraz czy jest zainstalowana biblioteka shapely.")
                return
            self._rm_krok(0, "done", "%d plików VAL" % ok_val)
            self._rm_trzymaj()

            # ================= KROK 2: ROZLICZANIE ==========================
            self._rm_krok(1, "running", "Rozliczanie...")
            if _pary_obrebow is not None:
                pary, _bez_val = _pary_obrebow(ewid_dir, val_dir)
                wybrane = {x.stem: (v.stem if v else "") for x, v in pary}
            else:
                pary, wybrane = [], None
            self.update_status("Krok 2/4 — rozliczanie powierzchni...", "#0078D7")
            self.log("[R+M] KROK 2/4 — rozliczanie powierzchni "
                     "(połączonych obrębów: %d)." % len(pary))
            self.run_rozliczanie_thread(str(ewid_dir), str(val_dir), str(rozl_dir),
                                        False, False, wybrane)
            self._rm_trzymaj()
            rozl_ile = len(list(rozl_dir.glob("*_Rozliczone.xlsx")))
            self.log("[R+M] Krok 2 gotowy: rozliczonych obrębów = %d" % rozl_ile)
            if rozl_ile == 0:
                self._rm_krok(1, "error", "0 rozliczeń")
                messagebox.showerror(
                    "Rozliczanie + MIETEK",
                    "Nie powstał żaden plik rozliczenia.\n\n"
                    "Sprawdź, czy nazwy plików ewidencji i .VAL do siebie pasują.")
                return
            self._rm_krok(1, "done", "%d obrębów" % rozl_ile)

            # ================= KROK 3: MIETKI + KRZYŻÓWKI ===================
            self._rm_krok(2, "running", "Tworzenie mietków...")
            _pominiete_typy = []
            try:
                if self._pokaz_wybory_op(str(ewid_dir)) is None:
                    self._rm_krok(2, "error", "przerwano")
                    self.log("[R+M] Przerwano na wyborze typów właścicieli.")
                    return
                _w = getattr(self, "_op_wybory", None) or {}
                _pominiete_typy = sorted(k for k, v in _w.items() if v == "pomin")
            except Exception as e:                              # noqa: BLE001
                self.log("[R+M] Pomijam okno typów właścicieli: %s" % e)
            if _pominiete_typy:
                self.log("[R+M] UWAGA: dla typów właścicieli (%s) wybrano „usuń” "
                         "— ich działki NIE zostaną wpisane do mietka i nie będzie "
                         "ich w krzyżówkach. Rozliczenie w Excelu będzie przez to "
                         "większe niż powierzchnia w mietku."
                         % ", ".join(_pominiete_typy))

            self.update_status("Krok 3/4 — tworzenie mietków...", "#0078D7")
            self.log("[R+M] KROK 3/4 — tworzenie mietków z krzyżówkami.")
            base_dir = _rm_resource_path("pusty")
            if not Path(base_dir).exists():
                self._rm_krok(2, "error", "brak zasobu „pusty”")
                messagebox.showerror("Rozliczanie + MIETEK",
                                     "Nie znaleziono wbudowanego folderu „pusty” "
                                     "w plikach programu:\n%s" % base_dir)
                return
            names_list = sorted({p.stem for p in ewid_dir.iterdir()
                                 if p.is_file()
                                 and p.suffix.lower() in {".xls", ".xlsx"}})
            wsie_meta = {
                "WOJEW": u.get("wojew", ""), "POWIAT": u.get("powiat", ""),
                "STAN_NA": u.get("stan", ""), "OBOW_OD": u.get("obod", ""),
                "OBOW_DO": u.get("obdo", ""), "NR_WSI": u.get("nrws") or "1",
                "ROK_ZAL": u.get("rokz", ""),
            }
            self.run_tworzenie_mietkow_thread(base_dir, str(mietki_dir), names_list,
                                              str(ewid_dir), str(rozl_dir),
                                              wsie_meta, True)
            self._rm_trzymaj()
            obreby = [d for d in mietki_dir.iterdir() if d.is_dir()]
            self.log("[R+M] Krok 3 gotowy: utworzonych obrębów = %d" % len(obreby))
            if not obreby:
                self._rm_krok(2, "error", "brak obrębów")
                messagebox.showerror("Rozliczanie + MIETEK",
                                     "Nie powstał żaden folder obrębu.")
                return
            self._rm_krok(2, "done", "%d obrębów" % len(obreby))

            # ================= KROK 4: ZESTAWIENIE ==========================
            if u.get("zestaw"):
                self._rm_krok(3, "running", "Zestawienie...")
                self.update_status("Krok 4/4 — zestawienie zbiorcze...", "#0078D7")
                self.log("[R+M] KROK 4/4 — zestawienie zbiorcze "
                         "(zapis w folderze głównym, obok podfolderu ROZLICZONE).")
                # plik zestawienia ma leżeć o poziom WYŻEJ — w folderze
                # głównym procesu, a nie w podfolderze ROZLICZONE
                self.run_zestawienie_thread(str(rozl_dir), out_folder=praca)
                self._rm_trzymaj()
                self._rm_krok(3, "done", "gotowe")
            else:
                self._rm_krok(3, "done", "pominięto")

            # ---------------- podsumowanie ----------------------------------
            self._rm_ustaw_stan(zakonczone=True, blad=False,
                                komunikat="Zakończono")
            self.update_status("Rozliczanie + MIETEK zakończone.", "#107C10",
                               animate=False)
            self.log("=" * 60)
            self.log("[R+M] ZAKOŃCZONO. Folder wynikowy: %s" % praca)
            self.log("[R+M]   MAPY: %d, VAL: %d, ROZLICZONE: %d, MIETKI: %d obrębów"
                     % (len(mapy), len(list(val_dir.glob("*"))),
                        len(list(rozl_dir.glob("*_Rozliczone.xlsx"))),
                        len([d for d in mietki_dir.iterdir() if d.is_dir()])))
            self.log("=" * 60)
            _zest = ("\n  • ZESTAWIENIE_ZBIORCZE.xlsx — zestawienie zbiorcze "
                     "(leży w folderze głównym)" if u.get("zestaw") else "")
            _ostrz = ("\n\nUWAGA: typy właścicieli (%s) miały ustawienie „usuń” — "
                      "ich działki nie trafiły do mietka, więc powierzchnia "
                      "rozliczona w Excelu jest większa niż w mietku."
                      % ", ".join(_pominiete_typy)) if _pominiete_typy else ""
            messagebox.showinfo(
                "Rozliczanie + MIETEK — zakończone",
                ("Gotowe!\n\nWszystko zapisane w folderze:\n%s\n\n"
                 "  • MAPY — kopie map wejściowych,\n"
                 "  • VAL — wygenerowane pliki .VAL,\n"
                 "  • EWIDENCJA — kopie plików ewidencji,\n"
                 "  • ROZLICZONE — pliki <WIEŚ>_Rozliczone.xlsx,\n"
                 "  • MIETKI — gotowe struktury obrębów z krzyżówkami."
                 % praca) + _zest + _ostrz)
        except InterruptedError:
            self.log("\n[R+M] ZADANIE PRZERWANE PRZEZ UŻYTKOWNIKA.")
            self.update_status("Przerwano", "#D83B01", animate=False)
        except Exception as e:                                  # noqa: BLE001
            self.log("\n[R+M] BŁĄD: %s" % e)
            self.log(traceback.format_exc())
            # przy błędzie też wskazujemy folder roboczy, żeby „Otwórz folder
            # wyników" nie otwierał czegoś zupełnie innego z poprzednich zadań
            try:
                _p = locals().get("praca")
                if _p and Path(_p).exists():
                    self.last_output_dir = _p
                    self._zapamietaj_folder_wynikow(_p)
            except Exception:                               # noqa: BLE001
                pass
            self._rm_ustaw_stan(zakonczone=True, blad=True,
                                komunikat="Błąd: %s" % e)
            self.update_status("Błąd: %s" % e, "#D83B01", animate=False)
        finally:
            # przywracamy prawdziwe zakończenie zadania
            try:
                if _orig_restore is not None:
                    self.restore_all_buttons = _orig_restore
            except Exception:                                   # noqa: BLE001
                pass
            self.running = False
            try:
                self.after(0, self.restore_all_buttons)
            except Exception:                                   # noqa: BLE001
                self.restore_all_buttons()
