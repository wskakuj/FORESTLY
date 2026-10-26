# -*- coding: utf-8 -*-
"""Testy diagnostyki błędów: bez Word/GUI, tylko biblioteka standardowa."""
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from app.core.diagnostyka_bledow import (
    opis_do_logu, zbuduj_raport, zapisz_raport_jsonl,
)


class TestDiagnostykaBledow(unittest.TestCase):
    def test_pokazuje_etap_plik_przyczyne_i_podpowiedz(self):
        error = FileNotFoundError(2, "No such file", "X:/test/OPTAX.TXT")
        r = zbuduj_raport(
            error, etap="TXT → PDF", zadanie="Pełny Automat")
        self.assertEqual(r["etap"], "TXT → PDF")
        self.assertEqual(r["plik"], "X:/test/OPTAX.TXT")
        self.assertEqual(r["typ"], "FileNotFoundError")
        self.assertIn("plik", r["wskazowka"].lower())
        self.assertIn("OPTAX.TXT", opis_do_logu(r))

    def test_uprawnienia_maja_wlasna_podpowiedz(self):
        r = zbuduj_raport(PermissionError("Access denied"), etap="Zapis Word",
                         plik=Path("C:/folder/REJESTR1.docx"))
        self.assertIn("Wordzie", r["wskazowka"])
        self.assertTrue(r["krytyczny"])

    def test_wpisy_jsonl_sa_dopisywane(self):
        with TemporaryDirectory() as tmp:
            r = zbuduj_raport(ValueError("zły format"), etap="DBF",
                             plik="O123.DBF", krytyczny=False)
            p = zapisz_raport_jsonl(r, tmp)
            zapisz_raport_jsonl(r, tmp)
            self.assertEqual(p.name, "bledy.jsonl")
            self.assertIn("_FORESTLY_DIAGNOSTYKA", p.parts)
            rows = [json.loads(t) for t in p.read_text(encoding="utf-8").splitlines()]
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0]["plik"], "O123.DBF")
            self.assertFalse(rows[0]["krytyczny"])

    def test_brak_folderu_nie_tworzy_pliku(self):
        r = zbuduj_raport(RuntimeError("problem"), etap="Start")
        self.assertIsNone(zapisz_raport_jsonl(r, None))

    def test_raport_ma_traceback(self):
        try:
            raise ValueError("niepoprawna liczba")
        except ValueError as exc:
            r = zbuduj_raport(exc, etap="Czytanie powierzchni")
        self.assertIn("ValueError", r["szczegoly"])
        self.assertIn("niepoprawna liczba", r["szczegoly"])


if __name__ == "__main__":
    unittest.main()
