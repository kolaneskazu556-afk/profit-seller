#!/usr/bin/env python3
"""
Генератор СИНТЕТИЧЕСКИХ отчётов WB и Ozon для разработки и тестов сервиса селлеров.

Все данные выдуманы. Структура повторяет реальные выгрузки, но названия колонок и
типы операций площадки меняют: перед продом сверьтесь с настоящими файлами.

Что создаётся в --out:
  wb_report_synthetic.xlsx          еженедельный детализированный отчёт WB (русские заголовки, как в ЛК)
  wb_report_synthetic.csv           то же, но с именами полей API (reportDetailByPeriod)
  ozon_accruals_synthetic.xlsx      отчёт о начислениях Ozon (как из ЛК)
  ozon_transactions_synthetic.json  операции в формате /v3/finance/transaction/list
  costs_synthetic.xlsx / .csv       таблица себестоимости (вход вашего сервиса)
  ground_truth_synthetic.csv        эталон: прибыль по артикулам ДО налогов (для автотестов)
  ground_truth_unallocated.csv      статьи без артикула (хранение, приёмка, размещение, часть рекламы)

Допущения эталона (можно менять под свою логику):
  * прибыль = выплата (с учётом комиссии, эквайринга, логистики) - штрафы - реклама по артикулу
              - (себестоимость + упаковка) * (продано - возвращено)
  * вернувшийся товар считается вернувшимся на склад целым
  * хранение, приёмка, удержания WB, размещение Ozon в прибыль по артикулам НЕ входят
  * в WB ppvz_for_pay уже содержит вычет эквайринга (проверьте на реальных данных)

Запуск:
  python generate_synthetic_reports.py --articles 40 --seed 42 --out synthetic_data
  python generate_synthetic_reports.py --dirty      # добавить «грязные» случаи
"""
import argparse
import csv
import json
import random
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

WB_COLS = [
    ("realizationreport_id", "Номер отчета"),
    ("rr_dt", "Дата конца отчетного периода"),
    ("nm_id", "Код номенклатуры"),
    ("sa_name", "Артикул поставщика"),
    ("ts_name", "Размер"),
    ("barcode", "Баркод"),
    ("supplier_oper_name", "Обоснование для оплаты"),
    ("doc_type_name", "Тип документа"),
    ("quantity", "Кол-во"),
    ("retail_price_withdisc_rub", "Цена розничная с учетом согласованной скидки"),
    ("retail_amount", "Вайлдберриз реализовал Товар (Пр)"),
    ("commission_percent", "Размер комиссии, %"),
    ("ppvz_for_pay", "К перечислению Продавцу за реализованный Товар"),
    ("delivery_rub", "Услуги по доставке товара покупателю"),
    ("acquiring_fee", "Эквайринг/Комиссии за организацию платежей"),
    ("penalty", "Общая сумма штрафов"),
    ("storage_fee", "Хранение"),
    ("deduction", "Удержания"),
    ("acceptance", "Операции на приемке"),
    ("order_dt", "Дата заказа покупателем"),
    ("sale_dt", "Дата продажи"),
    ("srid", "Srid"),
]
OZON_COLS = ["Группа услуг", "Тип начисления", "Артикул", "SKU", "Название товара",
             "Дата начисления", "Номер отправления", "Количество", "Цена продавца", "Сумма итого"]

KINDS = [("Футболка", 14), ("Худи", 17), ("Джинсы", 18), ("Кроссовки", 20),
         ("Рюкзак", 16), ("Кружка", 13), ("Чехол для телефона", 15), ("Свеча ароматическая", 14)]
COLORS = ["красная", "синяя", "зелёная", "чёрная", "белая", "серая"]
SIZES = ["S", "M", "L", "XL", "ONE"]


def r2(x):
    return round(x + 0.0, 2)


def gen_catalog(rng, n, loss_share):
    loss_idx = set(rng.sample(range(n), max(1, round(n * loss_share))))
    items = []
    for i in range(1, n + 1):
        kind, comm = rng.choice(KINDS)
        cost = r2(rng.uniform(120, 1400))
        loss = (i - 1) in loss_idx
        mult = rng.uniform(1.12, 1.3) if loss else rng.uniform(1.9, 3.2)
        size = rng.choice(SIZES)
        items.append(dict(
            article=f"ART-{i:03d}",
            nm_id=100000000 + i,
            sku=1100000000 + i,
            barcode=str(2000000000000 + i * 7),
            name=f"{kind} {rng.choice(COLORS)} {size}",
            size=size,
            cost=cost,
            packaging=rng.choice([15.0, 25.0]),
            price=float(round(cost * mult / 10) * 10),
            wb_comm=comm + rng.choice([4.5, 5.5, 6.5]),
            oz_comm=float(comm),
            return_rate=0.28 if loss else rng.choice([0.03, 0.05, 0.08]),
            heavy=kind in ("Кроссовки", "Рюкзак"),
            on_wb=rng.random() < 0.8,
            on_oz=rng.random() < 0.8,
            demand_wb=rng.choice([0.5, 1, 2, 4, 8]),
            demand_oz=rng.choice([0.5, 1, 2, 4, 8]),
            loss_planted=loss,
        ))
    return items


def new_truth():
    return defaultdict(lambda: dict(sold=0, returned=0, revenue=0.0, payout=0.0, penalties=0.0, ads=0.0))


# ------------------------------------------------------------------ WB
def gen_wb(rng, items, weeks, truth, unalloc):
    rep_ids = [900000001 + k for k in range(len(weeks))]

    def rep_idx(d):
        return min(max((d - weeks[0]).days // 7, 0), len(weeks) - 1)

    rows = []

    def row(d, it, oper, doc="", **kw):
        k = rep_idx(d)
        base = {c: 0 for c, _ in WB_COLS}
        base.update(realizationreport_id=rep_ids[k], rr_dt=str(weeks[k] + timedelta(6)),
                    nm_id=it["nm_id"] if it else 0, sa_name=it["article"] if it else "",
                    ts_name=it["size"] if it else "", barcode=it["barcode"] if it else "",
                    supplier_oper_name=oper, doc_type_name=doc, order_dt="", sale_dt=str(d), srid="")
        base.update(kw)
        rows.append(base)

    for ws in weeks:
        for it in items:
            if not it["on_wb"]:
                continue
            n = max(0, round(rng.gauss(it["demand_wb"], it["demand_wb"] * 0.5)))
            key = ("WB", it["article"])
            for _ in range(n):
                sale_dt = ws + timedelta(rng.randint(0, 6))
                order_dt = sale_dt - timedelta(rng.randint(1, 4))
                srid = f"{rng.getrandbits(32):08x}.0.0"
                price = r2(it["price"] * rng.choice([1, 1, 1, 0.95, 0.9]))
                acq = r2(price * 0.019)
                pay = r2(price * (1 - it["wb_comm"] / 100) - acq)
                log = r2(rng.uniform(55, 130) + (90 if it["heavy"] else 0))
                row(sale_dt, it, "Продажа", "Продажа", quantity=1, retail_price_withdisc_rub=price,
                    retail_amount=price, commission_percent=it["wb_comm"], ppvz_for_pay=pay,
                    acquiring_fee=acq, order_dt=str(order_dt), srid=srid)
                row(sale_dt + timedelta(1), it, "Логистика", delivery_rub=log, order_dt=str(order_dt), srid=srid)
                t = truth[key]
                t["sold"] += 1
                t["revenue"] += price
                t["payout"] += pay - log
                if rng.random() < it["return_rate"]:
                    ret_dt = sale_dt + timedelta(rng.randint(3, 12))
                    if ret_dt >= weeks[-1] + timedelta(7):
                        continue  # возврат ещё в пути, в периоде его нет
                    rlog = r2(log * 0.9)
                    row(ret_dt, it, "Возврат", "Возврат", quantity=1, retail_price_withdisc_rub=price,
                        retail_amount=-price, commission_percent=it["wb_comm"], ppvz_for_pay=-pay,
                        acquiring_fee=-acq, order_dt=str(order_dt), srid=srid)
                    row(ret_dt, it, "Логистика", delivery_rub=rlog, order_dt=str(order_dt), srid=srid)
                    t["returned"] += 1
                    t["revenue"] -= price
                    t["payout"] += -pay - rlog

        k = rep_idx(ws)
        end = ws + timedelta(6)
        # штрафы по артикулам
        for it in rng.sample([i for i in items if i["on_wb"]], 3):
            amt = rng.choice([300.0, 500.0, 1000.0])
            row(end - timedelta(rng.randint(0, 5)), it, "Штраф", penalty=amt)
            truth[("WB", it["article"])]["penalties"] += amt
        # статьи без артикула
        for oper, field, amt in [("Хранение", "storage_fee", r2(rng.uniform(1500, 4000))),
                                 ("Удержание", "deduction", r2(rng.uniform(2000, 6000))),
                                 ("Платная приёмка", "acceptance", rng.choice([0.0, 960.0, 1440.0]))]:
            if amt:
                row(end, None, oper, **{field: amt})
                unalloc.append(("WB", oper, str(end), amt))
    rows.sort(key=lambda r: (r["realizationreport_id"], r["sale_dt"]))
    return rows


# ------------------------------------------------------------------ Ozon
def gen_ozon(rng, items, weeks, truth, unalloc):
    postings, ads = [], []
    end_all = weeks[-1] + timedelta(7)
    for ws in weeks:
        for it in items:
            if not it["on_oz"]:
                continue
            n = max(0, round(rng.gauss(it["demand_oz"], it["demand_oz"] * 0.5)))
            key = ("OZON", it["article"])
            for _ in range(n):
                d = ws + timedelta(rng.randint(0, 6))
                pn = f"{rng.randint(10**9, 10**10 - 1)}-{rng.randint(1, 9999):04d}-1"
                price = r2(it["price"] * rng.choice([1, 1, 1, 0.95]))
                comm = r2(price * it["oz_comm"] / 100)
                log = r2(40 + price * 0.035 + (60 if it["heavy"] else 0))
                last = r2(min(max(price * 0.055, 20), 500))
                postings.append(dict(kind="sale", it=it, date=d, pn=pn, price=price, comm=comm, log=log, last=last))
                t = truth[key]
                t["sold"] += 1
                t["revenue"] += price
                t["payout"] += price - comm - log - last
                if rng.random() < it["return_rate"]:
                    rd = d + timedelta(rng.randint(3, 12))
                    if rd >= end_all:
                        continue
                    rlog, rproc = r2(log * 0.8), 55.0
                    postings.append(dict(kind="return", it=it, date=rd, pn=pn, price=price, comm=comm,
                                         rlog=rlog, rproc=rproc))
                    t["returned"] += 1
                    t["revenue"] -= price
                    t["payout"] += -price + comm - rlog - rproc
            # реклама по артикулу
            if it["on_oz"] and rng.random() < 0.4:
                amt = r2(rng.uniform(300, 1800))
                ads.append(dict(it=it, date=ws + timedelta(6), amount=amt))
                truth[key]["ads"] += amt
        unalloc.append(("OZON", "Размещение товаров", str(ws + timedelta(6)), r2(rng.uniform(1500, 3500))))
    return postings, ads


def ozon_accrual_rows(postings, ads, unalloc, rng, dirty):
    rows = []

    def art(a):
        return a.lower() + " " if dirty and rng.random() < 0.03 else a

    def add(group, typ, it, d, pn, qty, price, total):
        rows.append([group, typ, art(it["article"]) if it else "", it["sku"] if it else "",
                     it["name"] if it else "", str(d), pn, qty, price, total])

    for p in sorted(postings, key=lambda x: x["date"]):
        it, d, pn = p["it"], p["date"], p["pn"]
        if p["kind"] == "sale":
            add("Продажи", "Выручка", it, d, pn, 1, p["price"], p["price"])
            add("Продажи", "Вознаграждение Ozon", it, d, pn, 1, "", -p["comm"])
            add("Услуги доставки", "Логистика", it, d + timedelta(1), pn, 1, "", -p["log"])
            add("Услуги доставки", "Последняя миля", it, d + timedelta(1), pn, 1, "", -p["last"])
        else:
            add("Возвраты", "Выручка (возврат)", it, d, pn, 1, p["price"], -p["price"])
            add("Возвраты", "Вознаграждение Ozon (возврат)", it, d, pn, 1, "", p["comm"])
            add("Возвраты", "Обработка возврата", it, d, pn, 1, "", -p["rproc"])
            add("Услуги доставки", "Логистика возврата", it, d, pn, 1, "", -p["rlog"])
    for a in ads:
        add("Услуги продвижения", "Оплата за клик", a["it"], a["date"], "", "", "", -a["amount"])
    for plat, name, d, amt in unalloc:
        if plat == "OZON":
            add("Прочие услуги", name, None, d, "", "", "", -amt)
    return rows


def ozon_transactions(postings, ads, unalloc):
    ops, oid = [], 10000000000
    for p in sorted(postings, key=lambda x: x["date"]):
        oid += 1
        it = p["it"]
        item = [{"name": it["name"], "sku": it["sku"]}]
        posting = {"delivery_schema": "FBO", "order_date": f"{p['date'] - timedelta(2)} 14:22:00",
                   "posting_number": p["pn"], "warehouse_id": 0}
        if p["kind"] == "sale":
            services = [{"name": "MarketplaceServiceItemDirectFlowLogistic", "price": -p["log"]},
                        {"name": "MarketplaceServiceItemDelivToCustomer", "price": -p["last"]}]
            amount = p["price"] - p["comm"] + sum(s["price"] for s in services)
            ops.append(dict(operation_id=oid, operation_type="OperationAgentDeliveredToCustomer",
                            operation_date=f"{p['date']} 00:00:00", operation_type_name="Доставка покупателю",
                            accruals_for_sale=p["price"], sale_commission=-p["comm"], delivery_charge=0,
                            return_delivery_charge=0, amount=r2(amount), type="orders",
                            posting=posting, items=item, services=services))
        else:
            services = [{"name": "MarketplaceServiceItemReturnFlowLogistic", "price": -p["rlog"]},
                        {"name": "MarketplaceServiceItemReturnProcessing", "price": -p["rproc"]}]
            amount = -p["price"] + p["comm"] + sum(s["price"] for s in services)
            ops.append(dict(operation_id=oid, operation_type="ClientReturnAgentOperation",
                            operation_date=f"{p['date']} 00:00:00", operation_type_name="Получение возврата",
                            accruals_for_sale=-p["price"], sale_commission=p["comm"], delivery_charge=0,
                            return_delivery_charge=0, amount=r2(amount), type="returns",
                            posting=posting, items=item, services=services))
    for a in ads:  # названия operation_type для не-заказных операций здесь иллюстративные
        oid += 1
        ops.append(dict(operation_id=oid, operation_type="OperationMarketplaceCostPerClick",
                        operation_date=f"{a['date']} 00:00:00", operation_type_name="Оплата за клик",
                        accruals_for_sale=0, sale_commission=0, delivery_charge=0, return_delivery_charge=0,
                        amount=-a["amount"], type="services",
                        posting={"delivery_schema": "", "order_date": "", "posting_number": "", "warehouse_id": 0},
                        items=[{"name": a["it"]["name"], "sku": a["it"]["sku"]}], services=[]))
    for plat, name, d, amt in unalloc:
        if plat == "OZON":
            oid += 1
            ops.append(dict(operation_id=oid, operation_type="OperationMarketplaceServiceStorage",
                            operation_date=f"{d} 00:00:00", operation_type_name=name, accruals_for_sale=0,
                            sale_commission=0, delivery_charge=0, return_delivery_charge=0, amount=-amt,
                            type="services", posting={"delivery_schema": "", "order_date": "",
                                                      "posting_number": "", "warehouse_id": 0},
                            items=[], services=[]))
    ops.sort(key=lambda o: o["operation_date"])
    return {"result": {"operations": ops, "page_count": 1, "row_count": len(ops)}}


# ------------------------------------------------------------------ вывод
def save_xlsx(df, path, sheet):
    df.to_excel(path, index=False, sheet_name=sheet)
    wb = load_workbook(path)
    ws = wb[sheet]
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font = Font(name="Arial")
    for i, col in enumerate(df.columns, 1):
        width = min(max(len(str(col)), *(len(str(v)) for v in df[col].head(200))) + 2, 45)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--articles", type=int, default=40)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--loss-share", type=float, default=0.15, help="доля заведомо убыточных артикулов")
    ap.add_argument("--start", default="2026-10-05", help="понедельник первой недели")
    ap.add_argument("--weeks", type=int, default=4)
    ap.add_argument("--dirty", action="store_true", help="добавить грязные случаи")
    ap.add_argument("--out", default="synthetic_data")
    a = ap.parse_args(argv)

    rng = random.Random(a.seed)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    start = date.fromisoformat(a.start)
    weeks = [start + timedelta(7 * k) for k in range(a.weeks)]
    items = gen_catalog(rng, a.articles, a.loss_share)
    truth, unalloc = new_truth(), []

    # WB
    wb_rows = gen_wb(rng, items, weeks, truth, unalloc)
    wb_df = pd.DataFrame(wb_rows, columns=[c for c, _ in WB_COLS])
    wb_df.to_csv(out / "wb_report_synthetic.csv", index=False, encoding="utf-8-sig")
    save_xlsx(wb_df.rename(columns=dict(WB_COLS)), out / "wb_report_synthetic.xlsx", "Отчет")

    # Ozon
    postings, ads = gen_ozon(rng, items, weeks, truth, unalloc)
    oz_rows = ozon_accrual_rows(postings, ads, unalloc, rng, a.dirty)
    save_xlsx(pd.DataFrame(oz_rows, columns=OZON_COLS), out / "ozon_accruals_synthetic.xlsx", "Начисления")
    with open(out / "ozon_transactions_synthetic.json", "w", encoding="utf-8") as f:
        json.dump(ozon_transactions(postings, ads, unalloc), f, ensure_ascii=False, indent=2)

    # себестоимость
    cost_rows = []
    for it in items:
        if it["on_wb"]:
            cost_rows.append(["WB", it["article"], it["nm_id"], it["cost"], it["packaging"], ""])
        if it["on_oz"]:
            cost_rows.append(["OZON", it["article"], it["sku"], it["cost"], it["packaging"], ""])
    if a.dirty and len(cost_rows) > 10:
        dropped = [cost_rows.pop(3), cost_rows.pop(7)]          # артикулы без себестоимости
        dup = list(cost_rows[1]); dup[3] = r2(dup[3] * 1.1)
        cost_rows.append(dup)                                   # дубль с другой ценой
        cost_rows[5][3] = str(cost_rows[5][3]).replace(".", ",")  # десятичная запятая
        print("dirty: без себестоимости:", [d[:2] for d in dropped], "| дубль:", dup[:2])
    cost_df = pd.DataFrame(cost_rows, columns=["platform", "article", "sku_or_nm_id", "cost_price",
                                               "packaging", "vat_rate"])
    cost_df.to_csv(out / "costs_synthetic.csv", index=False, encoding="utf-8-sig")
    save_xlsx(cost_df, out / "costs_synthetic.xlsx", "Себестоимость")

    # эталон
    by_art = {i["article"]: i for i in items}
    gt = []
    for (plat, art), t in sorted(truth.items()):
        it = by_art[art]
        net_units = t["sold"] - t["returned"]
        cogs = r2((it["cost"] + it["packaging"]) * net_units)
        profit = r2(t["payout"] - t["penalties"] - t["ads"] - cogs)
        gt.append(dict(platform=plat, article=art, units_sold=t["sold"], units_returned=t["returned"],
                       revenue_net=r2(t["revenue"]), payout_net=r2(t["payout"]), penalties=r2(t["penalties"]),
                       ads_direct=r2(t["ads"]), cogs=cogs, profit_before_tax=profit,
                       is_loss=profit < 0, planted_loss_maker=it["loss_planted"]))
    gt_df = pd.DataFrame(gt)
    gt_df.to_csv(out / "ground_truth_synthetic.csv", index=False, encoding="utf-8-sig")
    un_df = pd.DataFrame(unalloc, columns=["platform", "item", "date", "amount"])
    un_df.to_csv(out / "ground_truth_unallocated.csv", index=False, encoding="utf-8-sig")

    print(f"WB строк: {len(wb_df)} | Ozon строк: {len(oz_rows)} | эталон: {len(gt_df)} пар платформа/артикул")
    print(f"Убыточных по эталону: {int(gt_df.is_loss.sum())}, из них заложено: "
          f"{int((gt_df.is_loss & gt_df.planted_loss_maker).sum())}")
    print("Файлы в:", out.resolve())



def generate(out, articles=40, seed=42, loss_share=0.15, start="2026-10-05", weeks=4, dirty=False):
    """Программный вызов генератора (используется демо и тестами)."""
    argv = ["--articles", str(articles), "--seed", str(seed), "--loss-share", str(loss_share),
            "--start", start, "--weeks", str(weeks), "--out", str(out)]
    if dirty:
        argv.append("--dirty")
    main(argv)


if __name__ == "__main__":
    main()
