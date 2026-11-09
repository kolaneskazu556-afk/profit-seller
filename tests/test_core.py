# -*- coding: utf-8 -*-
"""Тесты ядра на синтетических данных. Запуск: python -m unittest discover -s tests -v"""
import json
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import core  # noqa: E402
import synthetic  # noqa: E402
import contextlib, io  # noqa: E402


def read(p):
    return Path(p).read_bytes()


class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.d = Path(cls.tmp.name)
        with contextlib.redirect_stdout(io.StringIO()):
            synthetic.generate(cls.d / "clean", articles=40, seed=42)
            synthetic.generate(cls.d / "dirty", articles=40, seed=7, dirty=True)
        c = cls.d / "clean"
        cls.truth = pd.read_csv(c / "ground_truth_synthetic.csv")
        cls.costs = core.parse_costs("costs_synthetic.csv", read(c / "costs_synthetic.csv"), [])

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def run_all(self, files, folder="clean", settings=None):
        W, parsed = [], []
        for f in files:
            parsed.append(core.parse_report(f, read(self.d / folder / f), W))
        ops = core.combine_reports(parsed, W)
        return core.analyze(ops, self.costs, settings), W

    def compare(self, res, platform):
        got = {(a["platform"], a["article"]): a for a in res["articles"]}
        t = self.truth[self.truth.platform == platform]
        self.assertGreater(len(t), 5)
        for r in t.itertuples():
            a = got[(r.platform, r.article)]
            self.assertAlmostEqual(a["profit_direct"], r.profit_before_tax, delta=0.05, msg=f"{r.platform} {r.article}")
            self.assertEqual(int(a["qty_sold"]), r.units_sold)
            self.assertEqual(int(a["qty_returned"]), r.units_returned)
            self.assertAlmostEqual(a["revenue"], r.revenue_net, delta=0.05)
            self.assertAlmostEqual(a["payout_net"], r.payout_net, delta=0.05)
            self.assertEqual(a["status"] == "loss", bool(r.is_loss), msg=f"{r.article} {a['profit_direct']}")


class TestAgainstGroundTruth(Base):
    def test_wb_xlsx(self):
        res, _ = self.run_all(["wb_report_synthetic.xlsx"])
        self.compare(res, "WB")

    def test_wb_csv_api_names_equal_xlsx(self):
        a, _ = self.run_all(["wb_report_synthetic.xlsx"])
        b, _ = self.run_all(["wb_report_synthetic.csv"])
        self.assertAlmostEqual(a["totals"]["profit_before_tax"], b["totals"]["profit_before_tax"], places=2)

    def test_ozon_xlsx(self):
        res, _ = self.run_all(["ozon_accruals_synthetic.xlsx"])
        self.compare(res, "OZON")

    def test_ozon_json_matches_xlsx(self):
        a, _ = self.run_all(["ozon_accruals_synthetic.xlsx"])
        b, _ = self.run_all(["ozon_transactions_synthetic.json"])
        self.compare(b, "OZON")
        self.assertAlmostEqual(a["totals"]["profit_before_tax"], b["totals"]["profit_before_tax"], places=1)

    def test_both_platforms_and_unallocated(self):
        res, _ = self.run_all(["wb_report_synthetic.xlsx", "ozon_accruals_synthetic.xlsx"])
        un = pd.read_csv(self.d / "clean" / "ground_truth_unallocated.csv")
        self.assertAlmostEqual(res["totals"]["unallocated"], un.amount.sum(), delta=0.1)
        expected = self.truth.profit_before_tax.sum() - un.amount.sum()
        self.assertAlmostEqual(res["totals"]["profit_before_tax"], expected, delta=0.5)
        self.assertTrue(any(a["explain"] for a in res["articles"] if a["status"] == "loss"))

    def test_xlsx_and_json_not_double_counted(self):
        res, W = self.run_all(["ozon_accruals_synthetic.xlsx", "ozon_transactions_synthetic.json"])
        self.assertTrue(any("дважды" in w[1] for w in W))
        self.compare(res, "OZON")

    def test_taxes(self):
        r6, _ = self.run_all(["wb_report_synthetic.xlsx"], settings={"tax_mode": "usn6"})
        self.assertAlmostEqual(r6["tax"]["tax"], r6["totals"]["revenue"] * 0.06, delta=0.1)
        r15, _ = self.run_all(["wb_report_synthetic.xlsx"], settings={"tax_mode": "usn15"})
        self.assertGreaterEqual(r15["tax"]["tax"], 0)
        rn, _ = self.run_all(["wb_report_synthetic.xlsx"], settings={"tax_mode": "npd"})
        self.assertTrue(any("собственного производства" in n for n in rn["tax"]["notes"]))
        r0, _ = self.run_all(["wb_report_synthetic.xlsx"], settings={"tax_mode": "none"})
        self.assertEqual(r0["tax"]["tax"], 0)


class TestRobustness(Base):
    def test_missing_costs_flagged(self):
        W = []
        parsed = [core.parse_report("w.xlsx", read(self.d / "clean" / "wb_report_synthetic.xlsx"), W)]
        ops = core.combine_reports(parsed, W)
        res = core.analyze(ops, None)
        self.assertTrue(any("себестоимост" in w["text"] for w in res["warnings"]))
        self.assertTrue(any(a["status"] == "nocost" for a in res["articles"]))

    def test_dirty_data(self):
        W = []
        costs = core.parse_costs("c.csv", read(self.d / "dirty" / "costs_synthetic.csv"), W)
        self.assertTrue(any("разные цены" in w[1] for w in W))
        parsed = [core.parse_report("o.xlsx", read(self.d / "dirty" / "ozon_accruals_synthetic.xlsx"), W)]
        ops = core.combine_reports(parsed, W)
        self.assertFalse(ops["article"].str.contains(r"[a-zа-я]|\s", regex=True).any())  # нормализация регистра и пробелов
        res = core.analyze(ops, costs)
        self.assertTrue(any(a["status"] == "nocost" for a in res["articles"]))

    def test_garbage_and_empty(self):
        for name, data in [("a.csv", b""), ("b.csv", "x;y\n1;2\n".encode()), ("c.json", b"{bad"),
                           ("d.json", b'{"result":{"operations":[]}}'), ("e.pdf", b"%PDF")]:
            with self.assertRaises(core.ParseError, msg=name):
                core.parse_report(name, data, [])

    def test_header_after_title_rows_and_cp1251_semicolon(self):
        lines = ["Отчёт о начислениях;;;;;;;;;", ";;;;;;;;;", ";".join(core.OZ_ALIASES_HDR) if hasattr(core, "OZ_ALIASES_HDR") else
                 "Группа услуг;Тип начисления;Артикул;SKU;Название товара;Дата начисления;Номер отправления;Количество;Цена продавца;Сумма итого",
                 "Продажи;Выручка;ART-1;111;Тест;05.10.2026;P-1;1;1000,50;1000,50",
                 "Продажи;Вознаграждение Ozon;ART-1;111;Тест;05.10.2026;P-1;1;;-150,00",
                 "Услуги доставки;Логистика;ART-1;111;Тест;06.10.2026;P-1;1;;-70,25",
                 "Возвраты;Выручка (возврат);ART-1;111;Тест;10.10.2026;P-1;1;1000,50;-1000,50",
                 "Прочие услуги;Странная услуга;;;;11.10.2026;;;;-10,00"]
        W = []
        _, kind, df = core.parse_report("x.csv", "\n".join(lines).encode("cp1251"), W)
        self.assertEqual(kind, "ozon_xlsx")
        self.assertEqual(df["qty_sold"].sum(), 1)
        self.assertEqual(df["qty_returned"].sum(), 1)
        self.assertAlmostEqual(df["revenue"].sum(), 0.0)
        self.assertAlmostEqual(df["logistics"].sum(), 70.25)
        self.assertTrue(any("не распознан тип" in w[1] for w in W))
        self.assertEqual(df["date"].iloc[0], "2026-10-05")

    def test_negative_and_garbage_numbers(self):
        csv = ("Группа услуг;Тип начисления;Артикул;SKU;Дата начисления;Сумма итого\n"
               "Продажи;Выручка;A;1;2026-10-01;abc\nПродажи;Выручка;A;1;2026-10-01;−500,00\n"
               "Продажи;Выручка;A;1;2026-10-01;(100)\n")
        W = []
        _, _, df = core.parse_report("n.csv", csv.encode("utf-8"), W)
        self.assertTrue(any("не распознано" in w[1] for w in W))
        self.assertAlmostEqual(df["revenue"].sum(), -600.0)

    def test_costs_template_example_rows_ignored(self):
        W = []
        body = core.costs_template_csv() + "WB,ART-9,,100,5,0,\n"
        df = core.parse_costs("t.csv", body.encode("utf-8"), W)
        self.assertEqual(list(df["article"]), ["ART-9"])

    def test_json_amount_mismatch_warned(self):
        op = {"operation_type": "X", "operation_date": "2026-10-01 00:00:00", "accruals_for_sale": 100, "sale_commission": -10,
              "amount": 50, "items": [{"sku": 1, "name": "n"}], "services": [], "posting": {"posting_number": "p"}}
        W = []
        df = core.parse_ozon_json("t.json", json.dumps({"result": {"operations": [op]}}).encode(), W)
        self.assertTrue(any("не сходится" in w[1] for w in W))
        self.assertAlmostEqual(df["revenue"].sum() - df[core.COST_BUCKETS].sum().sum(), 50.0)


class TestPayouts(Base):
    def test_payout_reconciliation(self):
        W, parsed = [], [core.parse_report("w.xlsx", read(self.d / "clean" / "wb_report_synthetic.xlsx"), [])]
        ops = core.combine_reports(parsed, W)
        res = core.analyze(ops, self.costs)
        self.assertEqual(len(res["payouts"]), 4)
        p = res["payouts"][0]
        res2 = core.analyze(ops, self.costs, actuals={(p["platform"], p["key"]): p["expected"] - 100})
        self.assertAlmostEqual(res2["payouts"][0]["diff"], -100, places=1)
        self.assertAlmostEqual(sum(x["expected"] for x in res["payouts"]), res["totals"]["payout"], places=1)


if __name__ == "__main__":
    unittest.main()
