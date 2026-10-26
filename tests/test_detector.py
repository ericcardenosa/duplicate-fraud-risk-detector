"""Integration and edge-case tests for the Duplicate & Fraud Risk Detector.

Run from the project root:
    python -m unittest discover -s tests -v
"""
import csv
import random
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import fraud_detector as fd  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from scripts.exportar_seguimiento import export  # noqa: E402
from src.source_reader import Sources, load_data  # noqa: E402

INPUTS = ROOT / "inputs"


def invoice(row, num, vendor, total, date, desc="Servicio", refs=None, paid=0.0):
    return dict(num=num, row=row, date=date, acct_date=date, due=None, vendor=vendor, description=desc, tags="",
                account="", payment_method="", project="", subtotal=total, iva=0.0, withholding=0.0, total=total,
                paid=paid, outstanding=total - paid, state="", payment_date=None, refs=refs or [])


class DetectorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.src = load_data(INPUTS)
        cls.result = fd.detect(cls.src, days=7, similarity=0.86, threshold=1.0)

    def test_source_volumes_and_case_totals(self):
        self.assertEqual(len(self.src.invoices), 2239)
        self.assertEqual(len(self.src.orders), 427)
        cases = self.result["cases"]
        self.assertEqual(len(cases), 35)
        self.assertEqual(sum(c["type"] == "PAR" for c in cases), 11)
        self.assertEqual(sum(c["severity"] == "Crítica" for c in cases), 8)
        self.assertEqual(sum(c["score"] >= 72 for c in cases), 34)

    def test_every_rule_triggers_as_designed(self):
        per_rule = {code: 0 for code in fd.RULES}
        for c in self.result["cases"]:
            per_rule[c["rule"]] += 1
        self.assertEqual(per_rule, {"R01": 7, "R02": 1, "R03": 1, "R04": 1, "R05": 1, "R06": 24, "R07": 0})

    def test_case_ids_are_unique_and_independent_of_input_order(self):
        ids = [c["id"] for c in self.result["cases"]]
        self.assertEqual(len(ids), len(set(ids)))
        shuffled = Sources(list(self.src.orders), list(self.src.invoices), set(self.src.project_ids),
                           dict(self.src.cecos), dict(self.src.files))
        random.Random(3).shuffle(shuffled.invoices)
        random.Random(5).shuffle(shuffled.orders)
        again = fd.detect(shuffled, days=7, similarity=0.86, threshold=1.0)
        self.assertEqual([c["id"] for c in again["cases"]], ids)

    def test_credit_note_is_not_paired_with_its_invoice(self):
        d = datetime(2026, 3, 1)
        src = Sources(orders=[], invoices=[invoice(6, "FX26-1", "Proveedor A", 1000.0, d),
                                           invoice(7, "FX26-1", "Proveedor A", -1000.0, d)],
                      project_ids=set(), cecos={}, files={})
        self.assertEqual(fd.detect(src)["cases"], [])

    def test_monthly_series_ties_to_cases(self):
        dated = sum(1 for c in self.result["cases"] if c["date"] and c["date"].year == 2026)
        self.assertEqual(sum(row[2] for row in self.result["monthly"]), dated)

    def test_invalid_parameters_are_rejected(self):
        with self.assertRaises(ValueError):
            fd.detect(self.src, days=-1)
        with self.assertRaises(ValueError):
            fd.detect(self.src, similarity=1.5)

    def test_review_tracking_survives_regeneration(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "output" / "DUPLICATE_RISK_2026.xlsx"
            fd.build_excel(self.result, out, INPUTS)
            wb = load_workbook(out)
            self.assertEqual(len(wb.sheetnames), 12)
            ws = wb["02_Casos"]
            first_id = ws["A6"].value
            ws["N6"], ws["O6"], ws["Q6"] = "Justificada", "Revisor", "Servicios distintos"
            wb.save(out)
            csv_path = Path(tmp) / "seguimiento_casos.csv"
            self.assertEqual(export(out, csv_path), 1)
            review = fd.read_review_csv(csv_path)
            again = fd.detect(self.src, review=review)
            case = next(c for c in again["cases"] if c["id"] == first_id)
            self.assertEqual((case["status"], case["owner"], case["notes"]), ("Justificada", "Revisor", "Servicios distintos"))
            with csv_path.open(encoding="utf-8-sig") as fh:
                self.assertEqual(next(csv.reader(fh)), ["case_id", "estado", "responsable", "notas", "evidencia"])


if __name__ == "__main__":
    unittest.main()
