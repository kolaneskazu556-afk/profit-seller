# -*- coding: utf-8 -*-
"""
core.py: разбор отчётов WB и Ozon, таблицы себестоимости, расчёт прибыли и ориентировочных налогов.
Без веб-зависимостей, чтобы можно было тестировать отдельно (tests/test_core.py).

ВАЖНО: названия колонок и типов операций заложены по СИНТЕТИЧЕСКИМ примерам (sample_data/),
собранным по описанию выгрузок. На реальных отчётах не проверено. Все соответствия
вынесены в словари WB_ALIASES / OZ_ALIASES / COST_ALIASES: при расхождении правятся они.

Единая модель операции (одна строка = одна операция), суммы в рублях:
  revenue      выручка (продажи со знаком +, возвраты со знаком -)
  commission, acquiring, logistics, return_fees, penalty, storage, ads, other  расходы (положительное число = расход)
  чистый результат строки = revenue - (сумма расходов)
"""
from __future__ import annotations

import io
import json
import re

import numpy as np
import pandas as pd

BUCKETS = ["revenue", "commission", "acquiring", "logistics", "return_fees",
           "penalty", "storage", "ads", "other"]
COST_BUCKETS = BUCKETS[1:]
LABELS = {
    "revenue": "Выручка (с учётом возвратов)",
    "commission": "Комиссия площадки",
    "acquiring": "Эквайринг",
    "logistics": "Логистика",
    "return_fees": "Обработка возвратов",
    "penalty": "Штрафы",
    "storage": "Хранение, приёмка, размещение",
    "ads": "Реклама и удержания",
    "other": "Прочие расходы",
}
OPS_COLS = ["platform", "date", "article", "ext_id", "name", "order_id", "op_type",
            "qty_sold", "qty_returned"] + BUCKETS + ["report_key", "src_file", "src_row"]

MAX_ROWS = 600_000


class ParseError(Exception):
    """Ошибка разбора файла: текст можно показывать пользователю."""


# ---------------------------------------------------------------- нормализация
def norm(s) -> str:
    s = str(s).lower().replace("ё", "е").replace("\xa0", " ")
    s = re.sub(r"[^\w%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _lookup(aliases: dict) -> dict:
    return {norm(a): k for k, v in aliases.items() for a in v}


# Имена полей API (латиница) и заголовки xlsx из ЛК (русские). Правятся здесь.
WB_ALIASES = {
    "report_id": ["realizationreport_id", "Номер отчета", "Номер отчёта"],
    "rr_dt": ["rr_dt", "Дата конца отчетного периода", "Дата конца отчётного периода"],
    "nm_id": ["nm_id", "Код номенклатуры"],
    "article": ["sa_name", "Артикул поставщика"],
    "oper": ["supplier_oper_name", "Обоснование для оплаты"],
    "doc_type": ["doc_type_name", "Тип документа"],
    "qty": ["quantity", "Кол-во", "Количество"],
    "retail_amount": ["retail_amount", "Вайлдберриз реализовал Товар (Пр)"],
    "pay": ["ppvz_for_pay", "К перечислению Продавцу за реализованный Товар"],
    "delivery": ["delivery_rub", "Услуги по доставке товара покупателю"],
    "acquiring": ["acquiring_fee", "Эквайринг/Комиссии за организацию платежей"],
    "penalty": ["penalty", "Общая сумма штрафов"],
    "storage": ["storage_fee", "Хранение"],
    "deduction": ["deduction", "Удержания"],
    "acceptance": ["acceptance", "Операции на приемке"],
    "sale_dt": ["sale_dt", "Дата продажи"],
    "srid": ["srid", "Srid"],
}
OZ_ALIASES = {
    "group": ["Группа услуг"],
    "type": ["Тип начисления"],
    "article": ["Артикул", "Артикул продавца", "offer_id"],
    "sku": ["SKU"],
    "name": ["Название товара", "Наименование товара"],
    "date": ["Дата начисления"],
    "posting": ["Номер отправления"],
    "qty": ["Количество"],
    "amount": ["Сумма итого"],
}
COST_ALIASES = {
    "platform": ["platform", "Площадка", "Маркетплейс"],
    "article": ["article", "Артикул", "Артикул продавца"],
    "ext_id": ["sku_or_nm_id", "sku", "nm_id", "Код номенклатуры"],
    "cost": ["cost_price", "Себестоимость", "Закупочная цена", "Закупка"],
    "packaging": ["packaging", "Упаковка"],
    "other": ["other_cost", "Прочие расходы", "Прочее"],
    "vat": ["vat_rate", "Ставка НДС", "НДС"],
}
WB_LOOK, OZ_LOOK, COST_LOOK = _lookup(WB_ALIASES), _lookup(OZ_ALIASES), _lookup(COST_ALIASES)

PII_HINTS = ["фио", "телефон", "адрес", "email", "e mail", "покупател", "получател",
             "паспорт", "имя клиента", "фамилия"]


# ---------------------------------------------------------------- чтение файлов
def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def read_raw(name: str, data: bytes) -> pd.DataFrame:
    low = name.lower()
    if not data:
        raise ParseError(f"Файл «{name}» пустой.")
    if low.endswith((".xlsx", ".xlsm")):
        try:
            raw = pd.read_excel(io.BytesIO(data), header=None, dtype=object, engine="openpyxl")
        except Exception as e:  # noqa: BLE001
            raise ParseError(f"Не удалось прочитать Excel-файл «{name}»: {e}")
    elif low.endswith(".xls"):
        raise ParseError(f"Файл «{name}»: формат .xls не поддерживается, пересохраните в .xlsx или .csv.")
    elif low.endswith((".csv", ".txt", ".tsv")):
        text = _decode(data)
        head = "\n".join(text.splitlines()[:30])
        sep = max([";", ",", "\t"], key=head.count)
        try:
            raw = pd.read_csv(io.StringIO(text), sep=sep, header=None, dtype=str,
                              engine="python", on_bad_lines="skip", skip_blank_lines=True)
        except Exception as e:  # noqa: BLE001
            raise ParseError(f"Не удалось прочитать CSV «{name}»: {e}")
    else:
        raise ParseError(f"Файл «{name}»: поддерживаются .xlsx, .csv и .json.")
    raw = raw.dropna(how="all")
    if raw.empty:
        raise ParseError(f"Файл «{name}» не содержит данных.")
    if len(raw) > MAX_ROWS:
        raise ParseError(f"Файл «{name}» слишком большой ({len(raw)} строк, максимум {MAX_ROWS}).")
    return raw


def scan_header(raw: pd.DataFrame, lookup: dict, scan: int = 40):
    """Ищет строку заголовков среди первых строк: отчёты часто содержат шапку над таблицей."""
    best_row, best = -1, {}
    for i in range(min(scan, len(raw))):
        cols = {}
        for pos, v in enumerate(raw.iloc[i].tolist()):
            if pd.isna(v):
                continue
            k = lookup.get(norm(v))
            if k and k not in cols:
                cols[k] = pos
        if len(cols) > len(best):
            best_row, best = i, cols
    return best_row, best


def to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.astype("float64")
    t = s.astype("string").str.strip()
    t = (t.str.replace("\u2212", "-", regex=False).str.replace("\xa0", "", regex=False)
          .str.replace(" ", "", regex=False).str.replace(",", ".", regex=False))
    t = t.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    return pd.to_numeric(t, errors="coerce").astype("float64")


def to_date(s: pd.Series) -> pd.Series:
    st = s.astype("string").str.strip()
    a = pd.to_datetime(st.str[:10], format="%Y-%m-%d", errors="coerce")
    b = pd.to_datetime(st, dayfirst=True, errors="coerce", format="mixed")
    return a.where(a.notna(), b)


class _Tbl:
    """Доступ к колонкам таблицы по каноническим именам."""

    def __init__(self, raw, hdr, cols, fname, W):
        self.data = raw.iloc[hdr + 1:].dropna(how="all")
        self.cols, self.fname, self.W = cols, fname, W
        self.index = self.data.index

    def has(self, k):
        return k in self.cols

    def raw(self, k):
        return self.data.iloc[:, self.cols[k]] if k in self.cols else None

    def s(self, k) -> pd.Series:
        x = self.raw(k)
        if x is None:
            return pd.Series("", index=self.index, dtype="string")
        return x.astype("string").fillna("").str.strip()

    def num(self, k, label) -> pd.Series:
        x = self.raw(k)
        if x is None:
            return pd.Series(0.0, index=self.index)
        v = to_num(x)
        bad = int((v.isna() & (x.astype("string").fillna("").str.strip() != "")).sum())
        if bad:
            self.W.append(("warn", f"{self.fname}: в колонке «{label}» {bad} значений не распознано как число (приняты за 0)."))
        return v.fillna(0.0)

    def date(self, k) -> pd.Series:
        x = self.raw(k)
        if x is None:
            return pd.Series(pd.NaT, index=self.index, dtype="datetime64[ns]")
        return to_date(x)


def _pii_warn(raw, hdr, fname, W):
    names = [str(v) for v in raw.iloc[hdr].tolist() if not pd.isna(v)] if 0 <= hdr < len(raw) else []
    hits = [n for n in names if any(h in norm(n) for h in PII_HINTS)]
    if hits:
        W.append(("info", f"{fname}: найдены колонки с возможными данными покупателей ({', '.join(hits[:5])}). "
                          "Сервис читает только нужные для расчёта поля, остальные не сохраняются."))


def _finish(cols: dict, platform: str, fname: str) -> pd.DataFrame:
    df = pd.DataFrame(cols)
    for c in BUCKETS + ["qty_sold", "qty_returned"]:
        df[c] = df[c].astype("float64")
    keep = (df[BUCKETS].abs().sum(axis=1) > 0) | (df["qty_sold"] > 0) | (df["qty_returned"] > 0)
    df = df[keep].copy()
    df["platform"] = platform
    df["src_file"] = fname
    return df[OPS_COLS].reset_index(drop=True)


def _fmt_dates(d: pd.Series) -> np.ndarray:
    return d.dt.strftime("%Y-%m-%d").fillna("").to_numpy()


def _id_str(s: pd.Series) -> pd.Series:
    return s.str.replace(r"\.0+$", "", regex=True)


# ---------------------------------------------------------------- Wildberries
def parse_wb(raw, hdr, cols, fname, W) -> pd.DataFrame:
    t = _Tbl(raw, hdr, cols, fname, W)
    if t.data.empty:
        raise ParseError(f"«{fname}»: после строки заголовков нет данных.")
    for k, label in [("retail_amount", "Вайлдберриз реализовал Товар"), ("delivery", "Услуги по доставке"),
                     ("acquiring", "Эквайринг"), ("penalty", "Штрафы"), ("storage", "Хранение"),
                     ("deduction", "Удержания")]:
        if not t.has(k):
            W.append(("warn", f"{fname}: нет колонки «{label}»: эта статья не будет учтена в расчёте."))
    oper = t.s("oper").str.lower().str.replace("ё", "е", regex=False)
    doc = t.s("doc_type").str.lower()
    is_ret = (oper.str.contains("возврат", regex=False) | doc.str.contains("возврат", regex=False)).to_numpy(bool)
    is_sale = (oper.str.contains("продаж", regex=False)).to_numpy(bool) & ~is_ret
    low = oper.to_numpy(str)
    op_type = np.select(
        [is_ret, is_sale,
         np.char.find(low, "логист") >= 0, np.char.find(low, "штраф") >= 0,
         np.char.find(low, "хранен") >= 0, np.char.find(low, "приемк") >= 0,
         np.char.find(low, "удержан") >= 0],
        ["return", "sale", "logistics", "penalty", "storage", "acceptance", "deduction"], "other")
    retail, pay = t.num("retail_amount", "Вайлдберриз реализовал Товар").to_numpy(), t.num("pay", "К перечислению").to_numpy()
    acq = t.num("acquiring", "Эквайринг").to_numpy()
    deliv = t.num("delivery", "Услуги по доставке").to_numpy()
    pen = t.num("penalty", "Штрафы").to_numpy()
    stor = t.num("storage", "Хранение").to_numpy() + t.num("acceptance", "Операции на приемке").to_numpy()
    ded = t.num("deduction", "Удержания").to_numpy()
    deal = is_sale | is_ret
    unknown = (op_type == "other") & (np.abs(pay) > 0)
    qty = t.num("qty", "Кол-во").abs().to_numpy()
    qty = np.where(qty > 0, qty, 1.0)
    if unknown.any():
        names = pd.Series(low[unknown]).value_counts().head(5)
        W.append(("warn", f"{fname}: не распознан тип операции: " + "; ".join(f"«{k}» ({v} стр.)" for k, v in names.items())
                  + ". Суммы отнесены к «Прочие расходы»."))
    d = t.date("sale_dt")
    d = d.where(d.notna(), t.date("rr_dt"))
    if d.isna().mean() > 0.2:
        W.append(("warn", f"{fname}: у {d.isna().mean():.0%} строк не определена дата."))
    nm = t.num("nm_id", "Код номенклатуры").round().astype("int64")
    ext = np.where(nm.to_numpy() != 0, nm.astype(str).to_numpy(), "")
    rid = t.s("report_id")
    rkey = np.where(rid.to_numpy(str) != "", _id_str(rid).to_numpy(str), _fmt_dates(t.date("rr_dt")))
    _pii_warn(raw, hdr, fname, W)
    return _finish(dict(
        date=_fmt_dates(d), article=t.s("article").str.upper().to_numpy(str), ext_id=ext,
        name="", order_id=t.s("srid").to_numpy(str), op_type=op_type,
        qty_sold=np.where(is_sale, qty, 0.0), qty_returned=np.where(is_ret, qty, 0.0),
        revenue=retail, commission=np.where(deal, retail - pay - acq, 0.0), acquiring=acq,
        logistics=deliv, return_fees=0.0, penalty=pen, storage=stor, ads=ded,
        other=np.where(deal, 0.0, -pay), report_key=rkey, src_row=(t.index.to_numpy() + 1)),
        "WB", fname)


# ---------------------------------------------------------------- Ozon
def oz_bucket_text(text: str) -> str:
    """Тип начисления / название услуги (рус. или англ.) -> статья расходов."""
    t = norm(text)
    has = lambda *ks: any(k in t for k in ks)  # noqa: E731
    if has("вознагражд", "комисс", "commission"):
        return "commission"
    if has("обработк") or ("return" in t and has("process", "handl")):
        return "return_fees"
    if has("выручк", "accrualsforsale"):
        return "revenue"
    if has("логист", "последняя миля", "доставк", "logistic", "deliv"):
        return "logistics"
    if has("клик", "продвиж", "реклам", "трафарет", "cpc", "click", "promo", "advert", "stencil"):
        return "ads"
    if has("размещен", "хранен", "storage", "placement"):
        return "storage"
    if has("штраф", "penalt", "fine"):
        return "penalty"
    return "other"


def parse_ozon_xlsx(raw, hdr, cols, fname, W) -> pd.DataFrame:
    t = _Tbl(raw, hdr, cols, fname, W)
    if t.data.empty:
        raise ParseError(f"«{fname}»: после строки заголовков нет данных.")
    texts = (t.s("group") + " " + t.s("type")).str.strip()
    cache = {x: oz_bucket_text(x) for x in texts.unique()}
    bucket = np.array([cache[x] for x in texts], dtype=object)
    amt = t.num("amount", "Сумма итого").to_numpy()
    ret_text = texts.str.lower().str.contains("возврат", regex=False).to_numpy(bool)
    is_rev = bucket == "revenue"
    is_ret = is_rev & (ret_text | (amt < 0))
    is_sale = is_rev & ~is_ret
    qty = t.num("qty", "Количество").abs().to_numpy() if t.has("qty") else np.ones(len(amt))
    qty = np.where(qty > 0, qty, 1.0)
    out = {b: np.where(bucket == b, amt if b == "revenue" else -amt, 0.0) for b in BUCKETS}
    op_type = np.where(is_sale, "sale", np.where(is_ret, "return", bucket))
    oth = pd.DataFrame({"t": texts.to_numpy(str), "a": amt})[bucket == "other"]
    if len(oth):
        top = oth.groupby("t")["a"].agg(["size", "sum"]).sort_values("size", ascending=False).head(5)
        W.append(("warn", f"{fname}: не распознан тип начисления: "
                  + "; ".join(f"«{k}» ({int(r['size'])} стр., {r['sum']:,.0f} ₽)".replace(",", " ") for k, r in top.iterrows())
                  + ". Суммы отнесены к «Прочие расходы»."))
    d = t.date("date")
    if d.isna().mean() > 0.2:
        W.append(("warn", f"{fname}: у {d.isna().mean():.0%} строк не определена дата."))
    week = (d - pd.to_timedelta(d.dt.dayofweek, unit="D")).dt.strftime("%Y-%m-%d").fillna("")
    _pii_warn(raw, hdr, fname, W)
    return _finish(dict(
        date=_fmt_dates(d), article=t.s("article").str.upper().to_numpy(str),
        ext_id=_id_str(t.s("sku")).to_numpy(str), name=t.s("name").to_numpy(str),
        order_id=t.s("posting").to_numpy(str), op_type=op_type,
        qty_sold=np.where(is_sale, qty, 0.0), qty_returned=np.where(is_ret, qty, 0.0),
        report_key=week.to_numpy(str), src_row=(t.index.to_numpy() + 1), **out), "OZON", fname)


def parse_ozon_json(fname: str, data: bytes, W) -> pd.DataFrame:
    try:
        doc = json.loads(_decode(data))
    except Exception as e:  # noqa: BLE001
        raise ParseError(f"Файл «{fname}»: это не корректный JSON ({e}).")
    ops = doc.get("result", {}).get("operations") if isinstance(doc, dict) else doc
    if ops is None and isinstance(doc, dict):
        ops = doc.get("operations")
    if not isinstance(ops, list) or not ops:
        raise ParseError(f"Файл «{fname}»: не найден список операций (result.operations).")
    rows, mism = [], 0
    num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0  # noqa: E731
    for i, o in enumerate(ops):
        if not isinstance(o, dict):
            continue
        b = {k: 0.0 for k in BUCKETS}
        acc = num(o.get("accruals_for_sale"))
        amount = num(o.get("amount"))
        services = o.get("services") or []
        items = o.get("items") or []
        if acc != 0 or o.get("sale_commission") or services:
            b["revenue"] = acc
            b["commission"] += -num(o.get("sale_commission"))
            b["logistics"] += -(num(o.get("delivery_charge")) + num(o.get("return_delivery_charge")))
            for s in services:
                bk = oz_bucket_text(str(s.get("name", "")))
                bk = "other" if bk == "revenue" else bk
                b[bk] += -num(s.get("price"))
        else:
            bk = oz_bucket_text(f"{o.get('operation_type_name', '')} {o.get('operation_type', '')}")
            bk = "other" if bk == "revenue" else bk
            b[bk] += -amount
        net = b["revenue"] - sum(b[k] for k in COST_BUCKETS)
        if "amount" in o and abs(net - amount) > 0.01:
            mism += 1
            b["other"] += net - amount
        post = o.get("posting") or {}
        sku = items[0].get("sku") if items and isinstance(items[0], dict) else ""
        rows.append(dict(
            date=str(o.get("operation_date", ""))[:10], article="", ext_id=str(sku or ""),
            name=(items[0].get("name", "") if items and isinstance(items[0], dict) else ""),
            order_id=str(post.get("posting_number", "") if isinstance(post, dict) else ""),
            op_type="sale" if acc > 0 else "return" if acc < 0 else "service",
            qty_sold=float(max(1, len(items))) if acc > 0 else 0.0,
            qty_returned=float(max(1, len(items))) if acc < 0 else 0.0,
            report_key="", src_row=i + 1, **b))
    if mism:
        W.append(("warn", f"{fname}: у {mism} операций сумма amount не сходится с составляющими; разница отнесена к «Прочие расходы»."))
    df = pd.DataFrame(rows)
    if df.empty:
        raise ParseError(f"Файл «{fname}»: нет операций для разбора.")
    d = pd.to_datetime(df["date"], errors="coerce")
    df["report_key"] = (d - pd.to_timedelta(d.dt.dayofweek, unit="D")).dt.strftime("%Y-%m-%d").fillna("")
    return _finish({c: df[c].to_numpy() for c in df.columns}, "OZON", fname)


def parse_report(fname: str, data: bytes, W):
    """-> (platform, kind, ops DataFrame). kind: 'wb' | 'ozon_xlsx' | 'ozon_json'."""
    if fname.lower().endswith(".json"):
        return "OZON", "ozon_json", parse_ozon_json(fname, data, W)
    raw = read_raw(fname, data)
    rw, wcols = scan_header(raw, WB_LOOK)
    ro, ocols = scan_header(raw, OZ_LOOK)
    if {"oper", "pay", "article"} <= wcols.keys() and len(wcols) >= 5:
        return "WB", "wb", parse_wb(raw, rw, wcols, fname, W)
    if {"type", "amount", "date"} <= ocols.keys() and len(ocols) >= 4:
        return "OZON", "ozon_xlsx", parse_ozon_xlsx(raw, ro, ocols, fname, W)
    found = sorted(set(wcols) | set(ocols))
    raise ParseError(
        f"Файл «{fname}»: не удалось определить формат отчёта. Для WB нужны колонки «Обоснование для оплаты», "
        f"«Артикул поставщика», «К перечислению Продавцу…»; для Ozon: «Тип начисления», «Сумма итого», "
        f"«Дата начисления». Распознано полей: {', '.join(found) if found else 'нет'}. "
        "Возможно, площадка изменила названия колонок.")


def combine_reports(parsed: list, W) -> pd.DataFrame:
    """parsed: список (platform, kind, df). Убирает двойной учёт Ozon (xlsx + json)."""
    kinds = {k for _, k, _ in parsed}
    if "ozon_xlsx" in kinds and "ozon_json" in kinds:
        W.append(("warn", "Загружены и отчёт о начислениях Ozon (xlsx), и операции из API (json): "
                          "чтобы не считать дважды, использован xlsx, json пропущен."))
        parsed = [p for p in parsed if p[1] != "ozon_json"]
    if not parsed:
        return pd.DataFrame(columns=OPS_COLS)
    return pd.concat([p[2] for p in parsed], ignore_index=True)[OPS_COLS]


# ---------------------------------------------------------------- себестоимость
COST_COLS = ["platform", "article", "ext_id", "cost", "packaging", "other", "vat"]


def parse_costs(fname: str, data: bytes, W) -> pd.DataFrame:
    raw = read_raw(fname, data)
    hdr, cols = scan_header(raw, COST_LOOK)
    if "article" not in cols or "cost" not in cols:
        raise ParseError(f"Таблица себестоимости «{fname}»: нужны колонки «Артикул» (article) и «Себестоимость» (cost_price). "
                         f"Найдено: {', '.join(sorted(cols)) or 'ничего'}. Скачайте шаблон.")
    t = _Tbl(raw, hdr, cols, fname, W)
    plat = t.s("platform").str.lower().to_numpy(str)
    platform = np.where(np.char.find(plat, "wb") >= 0, "WB",
               np.where(np.char.find(plat, "wild") >= 0, "WB",
               np.where(np.char.find(plat, "вб") >= 0, "WB",
               np.where((np.char.find(plat, "ozon") >= 0) | (np.char.find(plat, "озон") >= 0), "OZON", "ALL"))))
    df = pd.DataFrame(dict(
        platform=platform, article=t.s("article").str.upper().to_numpy(str),
        ext_id=_id_str(t.s("ext_id")).to_numpy(str),
        cost=to_num(t.raw("cost")).to_numpy(), packaging=t.num("packaging", "Упаковка").to_numpy(),
        other=t.num("other", "Прочие расходы").to_numpy(),
        vat=to_num(t.raw("vat")).to_numpy() if t.has("vat") else np.nan))
    ex = df["article"].str.startswith("ПРИМЕР")
    if ex.any():
        df = df[~ex]
    empty = df["article"] == ""
    bad = df["cost"].isna()
    if (bad & ~empty).any():
        W.append(("warn", f"Таблица себестоимости: у {int((bad & ~empty).sum())} артикулов себестоимость не число, строки пропущены."))
    df = df[~empty & ~bad]
    neg = df["cost"] < 0
    if neg.any():
        W.append(("warn", f"Таблица себестоимости: у {int(neg.sum())} артикулов отрицательная себестоимость, строки пропущены."))
        df = df[~neg]
    dup = df.duplicated(["platform", "article"], keep=False)
    if dup.any():
        confl = df[dup].groupby(["platform", "article"])["cost"].nunique()
        confl = confl[confl > 1]
        if len(confl):
            W.append(("warn", "В таблице себестоимости у артикулов разные цены в нескольких строках: "
                      + ", ".join(f"{a} ({p})" for (p, a) in confl.index[:8]) + ". Взята первая строка."))
        df = df.drop_duplicates(["platform", "article"], keep="first")
    if df.empty:
        raise ParseError("Таблица себестоимости не содержит ни одной корректной строки.")
    return df[COST_COLS].reset_index(drop=True)


def costs_template_csv() -> str:
    return ("platform,article,sku_or_nm_id,cost_price,packaging,other_cost,vat_rate\n"
            "WB,ПРИМЕР-УДАЛИТЕ-1,100000001,420.00,15.00,0,\n"
            "OZON,ПРИМЕР-УДАЛИТЕ-2,1100000001,420.00,15.00,0,\n"
            "ALL,ПРИМЕР-УДАЛИТЕ-3,,300.00,10.00,0,\n")


# ---------------------------------------------------------------- расчёт
DEFAULT_SETTINGS = {"tax_mode": "usn6", "tax_rate": None, "income_base": "revenue",
                    "alloc": "revenue", "low_margin_pct": 10.0}
TAX_DEFAULT_RATE = {"none": 0.0, "usn6": 6.0, "usn15": 15.0, "npd": 4.0}
TAX_NAMES = {"none": "без расчёта", "usn6": "УСН «Доходы»", "usn15": "УСН «Доходы минус расходы»",
             "npd": "Самозанятый (НПД)"}


def _rub(x) -> str:
    return f"{x:,.0f}".replace(",", " ")


def _pct(x) -> str:
    return f"{x:.0f}%"


def expand_costs(costs: pd.DataFrame | None):
    if costs is None or costs.empty:
        return pd.DataFrame(columns=["platform", "article", "cost", "packaging", "other_cost"]), {}
    rows, idmap = [], {}
    for r in costs.itertuples(index=False):
        plats = ["WB", "OZON"] if r.platform == "ALL" else [r.platform]
        for p in plats:
            rows.append((p, r.article, r.cost, r.packaging, r.other, r.platform == "ALL"))
            if r.ext_id and (p, r.ext_id) not in idmap:
                idmap[(p, r.ext_id)] = r.article
    df = pd.DataFrame(rows, columns=["platform", "article", "cost", "packaging", "other_cost", "is_all"])
    df = df.sort_values("is_all", kind="stable").drop_duplicates(["platform", "article"], keep="first")
    return df.drop(columns="is_all"), idmap


def explain_article(r: dict, low_margin: float) -> str:
    sold, net_units = r["qty_sold"], r["qty_sold"] - r["qty_returned"]
    gross = r["sales_gross"]
    head = f"Артикул {r['article']} ({r['platform']})"
    if r["cost_missing"]:
        return f"{head}: нет себестоимости в таблице, прибыль посчитана без неё и может быть завышена."
    total = r["profit_direct"]
    unit = total / net_units if net_units > 0 else (total / sold if sold > 0 else None)
    if total < 0:
        lead = f"{head}: убыток {_rub(-total)} ₽ за период" + (f" (≈ {_rub(-unit)} ₽ с каждой проданной штуки)" if unit is not None else "")
    elif r["margin"] is not None and r["margin"] * 100 < low_margin:
        lead = f"{head}: маржа всего {_pct(r['margin'] * 100)} (прибыль {_rub(total)} ₽)"
    else:
        return f"{head}: прибыль {_rub(total)} ₽, маржа {_pct((r['margin'] or 0) * 100)}."
    causes = []
    if gross > 0:
        cogs_unit = (r["cogs"] / net_units) if net_units > 0 else 0
        items = [
            (r["cogs"] / gross, f"себестоимость и упаковка {_rub(cogs_unit)} ₽/шт. при средней цене {_rub(gross / max(sold, 1))} ₽ ({_pct(r['cogs'] / gross * 100)} цены)"),
            ((r["commission"] + r["acquiring"]) / gross, f"комиссия и эквайринг забирают {_pct((r['commission'] + r['acquiring']) / gross * 100)} выручки"),
            (r["logistics"] / gross, f"логистика {_pct(r['logistics'] / gross * 100)} выручки ({_rub(r['logistics'])} ₽)"),
            (r["ads"] / gross, f"реклама и удержания {_rub(r['ads'])} ₽ ({_pct(r['ads'] / gross * 100)} выручки)"),
            (r["penalty"] / gross, f"штрафы {_rub(r['penalty'])} ₽"),
        ]
        thr = {0: 0.45, 1: 0.30, 2: 0.15, 3: 0.12, 4: 0.03}
        for i, (share, text) in enumerate(items):
            if share >= thr[i]:
                causes.append((share, text))
    if sold > 0 and r["qty_returned"] / sold >= 0.15:
        causes.append((1.0, f"высокая доля возвратов: {r['qty_returned']:.0f} из {sold:.0f} шт. ({_pct(r['qty_returned'] / sold * 100)}); "
                            "за возвраты платите логистику, а выручка не остаётся"))
    causes.sort(key=lambda x: -x[0])
    if r["qty_sold"] == 0:
        causes = [(1, "продаж за период нет, но есть расходы (реклама, штрафы или хранение по этому артикулу)")]
    if not causes:
        causes = [(0, "суммарные расходы площадки и себестоимость почти равны цене продажи")]
    return lead + ". Основные причины: " + "; ".join(c[1] for c in causes[:3]) + "."


def analyze(ops: pd.DataFrame, costs: pd.DataFrame | None, settings: dict | None = None,
            actuals: dict | None = None) -> dict:
    st = {**DEFAULT_SETTINGS, **{k: v for k, v in (settings or {}).items() if v is not None or k == "tax_rate"}}
    W: list = []
    if ops is None or ops.empty:
        return {"empty": True, "warnings": [{"level": "info", "text": "Нет данных: загрузите отчёт площадки."}]}
    ops = ops.copy()
    for c in ("article", "ext_id", "name", "order_id", "report_key"):
        ops[c] = ops[c].fillna("").astype(str)
    for c in BUCKETS + ["qty_sold", "qty_returned"]:
        ops[c] = ops[c].astype(float)
    cost_x, idmap = expand_costs(costs)
    if costs is None or costs.empty:
        W.append(("warn", "Таблица себестоимости не загружена: прибыль показана без закупочных цен."))

    m = (ops["article"] == "") & (ops["ext_id"] != "")
    if m.any():
        new = [idmap.get((p, e)) or f"ID:{e}" for p, e in zip(ops.loc[m, "platform"], ops.loc[m, "ext_id"])]
        ops.loc[m, "article"] = new
        unres = sorted({n for n in new if n.startswith("ID:")})
        if unres:
            W.append(("warn", f"{len(unres)} товар(ов) в отчёте без артикула и без соответствия в таблице себестоимости (по SKU/nm_id): "
                      + ", ".join(unres[:6]) + ". Добавьте их в таблицу, колонка sku_or_nm_id."))

    ops["sales_gross"] = ops["revenue"].clip(lower=0)
    ops["refunds"] = (-ops["revenue"]).clip(lower=0)
    key = ["platform", "article"]
    art_ops = ops[ops["article"] != ""]
    agg = {c: "sum" for c in ["qty_sold", "qty_returned", "sales_gross", "refunds"] + BUCKETS}
    g = art_ops.groupby(key, as_index=False).agg(agg)
    names = art_ops[art_ops["name"] != ""].groupby(key)["name"].first().rename("name")
    g = g.merge(names, on=key, how="left")
    g["name"] = g["name"].fillna("")
    g = g.merge(cost_x, on=key, how="left")
    g["packaging"] = g["packaging"].fillna(0.0)
    g["other_cost"] = g["other_cost"].fillna(0.0)

    g["net_units"] = g["qty_sold"] - g["qty_returned"]
    g["cost_missing"] = g["cost"].isna() & ((g["qty_sold"] > 0) | (g["qty_returned"] > 0))
    g["cogs"] = (g["cost"].fillna(0.0) + g["packaging"] + g["other_cost"]) * g["net_units"]
    g["payout_net"] = g["revenue"] - g["commission"] - g["acquiring"] - g["logistics"] - g["return_fees"]
    g["direct_other"] = g["penalty"] + g["storage"] + g["ads"] + g["other"]
    g["profit_direct"] = g["payout_net"] - g["direct_other"] - g["cogs"]
    if (g["net_units"] < 0).any():
        W.append(("warn", f"У {int((g['net_units'] < 0).sum())} артикулов возвратов больше, чем продаж за период (продажа была до начала периода). "
                          "Себестоимость по ним посчитана со знаком «минус»."))

    # общие расходы без артикула
    un = ops[ops["article"] == ""]
    unalloc = {}
    for p in sorted(ops["platform"].unique()):
        u = un[un["platform"] == p]
        costs_u = float(u[COST_BUCKETS].sum().sum() - u["revenue"].sum())
        unalloc[p] = {"total": costs_u, **{b: float(u[b].sum()) for b in COST_BUCKETS}}
    method = st["alloc"]
    g["alloc"] = 0.0
    for p, u in unalloc.items():
        mk = g["platform"] == p
        basis = g.loc[mk, "sales_gross"] if method == "revenue" else g.loc[mk, "qty_sold"]
        if method != "none" and u["total"] != 0 and basis.sum() > 0:
            g.loc[mk, "alloc"] = u["total"] * basis / basis.sum()
        elif method != "none" and u["total"] != 0:
            W.append(("warn", f"{p}: общие расходы без артикула ({_rub(u['total'])} ₽) нельзя распределить: нет продаж."))
    g["profit_alloc"] = g["profit_direct"] - g["alloc"]
    g["margin"] = np.where(g["revenue"] > 0, g["profit_direct"] / g["revenue"].where(g["revenue"] > 0, 1), np.nan)
    g["return_rate"] = np.where(g["qty_sold"] > 0, g["qty_returned"] / g["qty_sold"].where(g["qty_sold"] > 0, 1) * 100, 0.0)
    low = float(st["low_margin_pct"])
    g["status"] = np.select(
        [g["cost_missing"], g["profit_direct"] < 0, (g["margin"] * 100) < low], ["nocost", "loss", "low"], "ok")
    g["loss_after_alloc"] = (g["status"].isin(["ok", "low"])) & (g["profit_alloc"] < 0)

    # итоги
    tot = {b: float(ops[b].sum()) for b in BUCKETS}
    cogs_total = float(g["cogs"].sum())
    unalloc_total = float(sum(u["total"] for u in unalloc.values()))
    payout_total = tot["revenue"] - sum(tot[b] for b in COST_BUCKETS)
    profit_before_tax = float(g["profit_direct"].sum()) - unalloc_total
    revenue_net = tot["revenue"]

    # налоги
    mode = st["tax_mode"] if st["tax_mode"] in TAX_DEFAULT_RATE else "usn6"
    rate = TAX_DEFAULT_RATE[mode] if st.get("tax_rate") in (None, "") else float(st["tax_rate"])
    expenses = sum(tot[b] for b in COST_BUCKETS) + cogs_total
    income = revenue_net if st["income_base"] == "revenue" else payout_total
    notes = [
        "Расчёт ориентировочный: не заменяет бухгалтера, не учитывает страховые взносы, вычеты, региональные ставки (если не заданы вручную), сроки уплаты и особенности учёта доходов у площадок.",
        "Ставки: УСН «Доходы» 6% и «Доходы минус расходы» 15% (регионы могут снижать; ставку можно изменить вручную); НПД 4% при продаже физлицам, 6% юрлицам и ИП. Проверено по открытым источникам на 05.11.2026, сверяйте с ФНС.",
        "С 2026 года на УСН при доходе за год свыше 20 млн ₽ возникает НДС (5% до 272,5 млн, 7% до 490,5 млн, либо общая 22%). НДС здесь НЕ рассчитывается.",
    ]
    if st["income_base"] == "revenue":
        notes.append("База дохода: выручка по отчётам (продажи минус возвраты) до вычета комиссий площадки. Как именно признавать доход при работе через маркетплейс, уточните у бухгалтера.")
    else:
        notes.append("База дохода: сумма к получению от площадки (после вычетов). Для УСН это упрощение, уточните у бухгалтера.")
    tax_total, min_tax = 0.0, None
    if mode == "usn6":
        tax_total = max(income, 0) * rate / 100
        notes.append("УСН 6%: налог не уменьшен на страховые взносы ИП.")
    elif mode == "usn15":
        base = max(income - expenses, 0)
        tax_total = base * rate / 100
        min_tax = max(income, 0) * 0.01
        notes.append(f"УСН 15%: расходы = удержания площадки + себестоимость проданного ({_rub(expenses)} ₽). Минимальный налог 1% от доходов ({_rub(min_tax)} ₽) сравнивается с налогом по итогам года, а не за период.")
    elif mode == "npd":
        tax_total = max(income, 0) * rate / 100
        notes.append("НПД: самозанятые вправе продавать только товары собственного производства, перепродажа покупных товаров не допускается. Годовой лимит дохода 2,4 млн ₽.")
    if mode != "none" and mode != "npd" and revenue_net * 12 > 20_000_000 and st["income_base"] == "revenue":
        notes.append("Выручка за период, пересчитанная на год, может превышать 20 млн ₽: проверьте, не возникла ли обязанность по НДС.")
    pos = g["profit_direct"].clip(lower=0)
    if mode in ("usn6", "npd"):
        b_i = g["revenue"].clip(lower=0) if st["income_base"] == "revenue" else (g["payout_net"] - g["direct_other"]).clip(lower=0)
        g["tax_est"] = b_i * rate / 100
    elif mode == "usn15" and pos.sum() > 0:
        g["tax_est"] = tax_total * pos / pos.sum()
    else:
        g["tax_est"] = 0.0
    g["profit_after_tax"] = g["profit_direct"] - g["tax_est"]

    # статьи для колонок / выводы
    arts = []
    for r in g.sort_values("profit_direct").to_dict("records"):
        r["explain"] = explain_article(r, low) if r["status"] in ("loss", "low", "nocost") else ""
        arts.append(r)
    loss = [a for a in arts if a["status"] == "loss"]
    lowm = [a for a in arts if a["status"] == "low"]
    nocost = [a for a in arts if a["status"] == "nocost"]
    if nocost:
        W.append(("warn", f"Нет себестоимости у {len(nocost)} артикулов с продажами (выручка {_rub(sum(a['sales_gross'] for a in nocost))} ₽): "
                  + ", ".join(a["article"] for a in nocost[:8]) + ". Прибыль по ним завышена, в список убыточных они не попадают."))

    dates = [d for d in ops["date"] if d]
    d1, d2 = (min(dates), max(dates)) if dates else ("", "")
    insights = [f"Период по данным: {d1} — {d2}. Выручка (продажи минус возвраты) {_rub(revenue_net)} ₽, прибыль до налога {_rub(profit_before_tax)} ₽"
                + (f", маржа {_pct(profit_before_tax / revenue_net * 100)}." if revenue_net > 0 else ".")]
    if loss:
        insights.append(f"Убыточных артикулов: {len(loss)} (суммарный убыток {_rub(-sum(a['profit_direct'] for a in loss))} ₽).")
        insights += [a["explain"] for a in loss[:3]]
    else:
        insights.append("Убыточных артикулов по прямым расходам не найдено.")
    if lowm:
        insights.append(f"Малоприбыльных артикулов (маржа ниже {low:.0f}%): {len(lowm)}.")
    la = [a for a in arts if a["loss_after_alloc"]]
    if la:
        insights.append(f"Ещё {len(la)} артикул(ов) уходят в минус, если распределить общие расходы (хранение, приёмка и т. п.): " + ", ".join(a["article"] for a in la[:6]) + ".")
    if revenue_net > 0 and unalloc_total:
        insights.append(f"Расходы без привязки к артикулу (хранение, приёмка, размещение, удержания): {_rub(unalloc_total)} ₽ ({_pct(unalloc_total / revenue_net * 100)} выручки).")
    if revenue_net > 0:
        big = max(COST_BUCKETS, key=lambda b: tot[b])
        if tot[big] > 0:
            insights.append(f"Крупнейшая статья расходов площадки: «{LABELS[big]}», {_rub(tot[big])} ₽ ({_pct(tot[big] / revenue_net * 100)} выручки).")
    sold_all = float(g["qty_sold"].sum())
    if sold_all > 0 and g["qty_returned"].sum() / sold_all >= 0.1:
        insights.append(f"Доля возвратов {_pct(g['qty_returned'].sum() / sold_all * 100)} от проданных штук: это заметно съедает прибыль.")

    # выплаты по отчётам
    ops["net"] = ops["revenue"] - ops[COST_BUCKETS].sum(axis=1)
    pg = ops[ops["report_key"] != ""].groupby(["platform", "report_key"]).agg(
        expected=("net", "sum"), d1=("date", lambda s: min([x for x in s if x] or [""])),
        d2=("date", lambda s: max([x for x in s if x] or [""])), rows=("net", "size")).reset_index()
    payouts = []
    for r in pg.sort_values(["platform", "report_key"]).to_dict("records"):
        label = (f"WB: отчёт №{r['report_key']}" if r["platform"] == "WB" else f"Ozon: неделя с {r['report_key']}")
        act = (actuals or {}).get((r["platform"], r["report_key"]))
        payouts.append(dict(platform=r["platform"], key=r["report_key"], label=label, date_from=r["d1"], date_to=r["d2"],
                            rows=int(r["rows"]), expected=r["expected"], actual=act,
                            diff=None if act is None else act - r["expected"]))
    pay_note = ("Показана сумма, которая должна прийти по данным отчёта (продажи минус все удержания). Даты перечисления определяет площадка: "
                "сервис их не прогнозирует. Введите фактически полученную сумму, чтобы увидеть расхождение. "
                "Для Ozon группировка по календарным неделям приблизительная.")

    method_notes = [
        "Прибыль по артикулу (прямая) = выручка − комиссия − эквайринг − логистика − обработка возвратов − штрафы − реклама − хранение по артикулу − (себестоимость + упаковка + прочее) × (продано − возвращено).",
        "Возвращённый товар считается вернувшимся на склад целым (без потерь).",
        "Расходы без артикула (хранение, приёмка, часть удержаний, размещение) в прямую прибыль не входят: они вычитаются в итоге и распределяются по артикулам "
        + {"revenue": "пропорционально выручке", "units": "пропорционально числу проданных штук", "none": "не распределяются"}[method] + " (колонка «С общими расходами»).",
        "Комиссия WB считается как разница: розничная сумма − «к перечислению» − эквайринг. Это расчётная величина, а не отдельная колонка отчёта.",
        "Названия колонок и типов операций заложены по синтетическим примерам и не проверены на реальных отчётах площадок.",
    ]
    out = dict(
        empty=False,
        meta=dict(period_from=d1, period_to=d2, rows=int(len(ops)), articles=int(len(g)),
                  platforms=sorted(ops["platform"].unique().tolist()), settings=dict(st, tax_mode=mode, tax_rate=rate),
                  files=sorted(ops["src_file"].unique().tolist())),
        totals=dict(**tot, cogs=cogs_total, unallocated=unalloc_total, payout=payout_total,
                    profit_before_tax=profit_before_tax, margin=(profit_before_tax / revenue_net if revenue_net > 0 else None),
                    tax=tax_total, profit_after_tax=profit_before_tax - tax_total),
        unallocated=unalloc,
        tax=dict(mode=mode, name=TAX_NAMES[mode], rate=rate, income=income, expenses=expenses, tax=tax_total,
                 min_tax=min_tax, notes=notes),
        articles=arts, insights=insights, payouts=payouts, payouts_note=pay_note, methodology=method_notes,
        labels=LABELS, warnings=[{"level": a, "text": b} for a, b in W])
    return clean(out)


def clean(o):
    """numpy/NaN -> обычные типы JSON, округление денег."""
    if isinstance(o, dict):
        return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [clean(v) for v in o]
    if isinstance(o, (np.bool_,)):
        return bool(o)
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (float, np.floating)):
        return None if (np.isnan(o) or np.isinf(o)) else round(float(o), 4 if abs(o) < 5 else 2)
    return o
