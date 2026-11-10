# -*- coding: utf-8 -*-
"""
Финансовый сервис для продавцов WB/Ozon: FastAPI + SQLite.
Запуск локально:  uvicorn app:app --reload
НЕ запускалось в среде разработки (там не было fastapi): ядро (core.py) покрыто тестами, этот файл нужно
проверить первым запуском и /ping.
"""
import csv
import hashlib
import io
import json
import os
import re
import sqlite3
import tempfile
import threading
import time
from collections import defaultdict, deque
from contextlib import closing
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, Header, HTTPException, Query, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse, Response

load_dotenv()
import core  # noqa: E402
import llm  # noqa: E402
import report_pdf  # noqa: E402
import synthetic  # noqa: E402

BASE = Path(__file__).resolve().parent
DB_PATH = os.getenv("DB_PATH", str(BASE / "data.db"))
RETENTION_DAYS = int(os.getenv("RETENTION_DAYS", "30"))
MAX_FILE = 20 * 1024 * 1024
MAX_FILES = 6
UID_RE = re.compile(r"^[A-Za-z0-9\-]{16,64}$")

app = FastAPI(title="Профит селлера", docs_url=None, redoc_url=None, openapi_url=None)
_lock = threading.Lock()
_hits = defaultdict(deque)


# ------------------------------------------------------------------ база
def db():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.execute("PRAGMA journal_mode=WAL")
    return con


def init_db():
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    cols = ", ".join(f"{c} {'REAL' if c in core.BUCKETS + ['qty_sold', 'qty_returned'] else 'TEXT'}" for c in core.OPS_COLS)
    with closing(db()) as con:
        con.executescript(f"""
        CREATE TABLE IF NOT EXISTS users(id TEXT PRIMARY KEY, created TEXT, last_seen TEXT);
        CREATE TABLE IF NOT EXISTS ops(user_id TEXT, {cols});
        CREATE INDEX IF NOT EXISTS ops_user ON ops(user_id);
        CREATE TABLE IF NOT EXISTS costs(user_id TEXT, platform TEXT, article TEXT, ext_id TEXT, cost REAL, packaging REAL, other REAL, vat REAL);
        CREATE INDEX IF NOT EXISTS costs_user ON costs(user_id);
        CREATE TABLE IF NOT EXISTS settings(user_id TEXT PRIMARY KEY, json TEXT);
        CREATE TABLE IF NOT EXISTS actuals(user_id TEXT, platform TEXT, key TEXT, amount REAL, PRIMARY KEY(user_id, platform, key));
        """)
        con.commit()


def wipe(con, uid):
    for t in ("ops", "costs", "settings", "actuals"):
        con.execute(f"DELETE FROM {t} WHERE user_id IN (?, ?)", (uid, "demo:" + uid))
    con.execute("DELETE FROM users WHERE id=?", (uid,))


def cleanup():
    lim = (datetime.utcnow() - timedelta(days=RETENTION_DAYS)).isoformat()
    with closing(db()) as con:
        for (uid,) in con.execute("SELECT id FROM users WHERE last_seen < ?", (lim,)).fetchall():
            wipe(con, uid)
        con.commit()


def _cleanup_loop():
    while True:
        try:
            cleanup()
        except Exception as e:  # noqa: BLE001
            print("cleanup error:", e)
        time.sleep(6 * 3600)


@app.on_event("startup")
def _startup():
    init_db()
    threading.Thread(target=_cleanup_loop, daemon=True).start()


def uid_of(x_user_id: str) -> str:
    if not UID_RE.match(x_user_id or ""):
        raise HTTPException(400, "Некорректный идентификатор пользователя. Обновите страницу.")
    now = datetime.utcnow().isoformat()
    with closing(db()) as con:
        con.execute("INSERT INTO users(id, created, last_seen) VALUES(?,?,?) ON CONFLICT(id) DO UPDATE SET last_seen=excluded.last_seen",
                    (x_user_id, now, now))
        con.commit()
    return x_user_id


def limit(key: str, n: int, per: int):
    q, now = _hits[key], time.time()
    while q and q[0] < now - per:
        q.popleft()
    if len(q) >= n:
        raise HTTPException(429, "Слишком много запросов, подождите немного.")
    q.append(now)


# ------------------------------------------------------------------ хранение и расчёт
def save(key: str, ops, costs, settings, replace_ops=True, replace_costs=True):
    with closing(db()) as con:
        if ops is not None:
            if replace_ops:
                con.execute("DELETE FROM ops WHERE user_id=?", (key,))
            o = ops.copy()
            o.insert(0, "user_id", key)
            o.to_sql("ops", con, if_exists="append", index=False)
        if costs is not None:
            if replace_costs:
                con.execute("DELETE FROM costs WHERE user_id=?", (key,))
            c = costs.copy()
            c.insert(0, "user_id", key)
            c.to_sql("costs", con, if_exists="append", index=False)
        if settings is not None:
            con.execute("INSERT INTO settings(user_id, json) VALUES(?,?) ON CONFLICT(user_id) DO UPDATE SET json=excluded.json",
                        (key, json.dumps(settings, ensure_ascii=False)))
        con.commit()


def load(key: str):
    with closing(db()) as con:
        ops = pd.read_sql_query(f"SELECT {', '.join(core.OPS_COLS)} FROM ops WHERE user_id=?", con, params=(key,))
        costs = pd.read_sql_query(f"SELECT {', '.join(core.COST_COLS)} FROM costs WHERE user_id=?", con, params=(key,))
        row = con.execute("SELECT json FROM settings WHERE user_id=?", (key,)).fetchone()
        act = {(p, k): a for p, k, a in con.execute("SELECT platform, key, amount FROM actuals WHERE user_id=?", (key,))}
    return ops, (costs if len(costs) else None), (json.loads(row[0]) if row else {}), act


def build(key: str, demo=False) -> dict:
    ops, costs, st, act = load(key)
    parse_w = st.pop("parse_warnings", [])
    res = core.analyze(ops, costs, st, act)
    if res.get("empty"):
        return res
    res["warnings"] = [{"level": a, "text": b} for a, b in parse_w] + res["warnings"]
    res["meta"]["demo"] = demo
    return res


def clean_settings(tax_mode, tax_rate, income_base, alloc, low=None) -> dict:
    s = {}
    if tax_mode in core.TAX_DEFAULT_RATE:
        s["tax_mode"] = tax_mode
    try:
        r = float(str(tax_rate).replace(",", ".")) if str(tax_rate).strip() else None
        s["tax_rate"] = r if r is not None and 0 <= r <= 30 else None
    except ValueError:
        s["tax_rate"] = None
    if income_base in ("revenue", "payout"):
        s["income_base"] = income_base
    if alloc in ("revenue", "units", "none"):
        s["alloc"] = alloc
    if low is not None:
        try:
            s["low_margin_pct"] = min(max(float(low), 0), 100)
        except (TypeError, ValueError):
            pass
    return s


def key_for(uid, demo):
    return "demo:" + uid if demo else uid


def read_upload(f: UploadFile) -> bytes:
    data = f.file.read(MAX_FILE + 1)
    if len(data) > MAX_FILE:
        raise HTTPException(413, f"Файл «{f.filename}» больше {MAX_FILE // 1024 // 1024} МБ.")
    return data


# ------------------------------------------------------------------ страницы и служебные адреса
@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse((BASE / "index.html").read_text(encoding="utf-8"),
                        headers={"Cache-Control": "no-cache", "X-Content-Type-Options": "nosniff"})


@app.get("/ping")
def ping():
    return PlainTextResponse("ok")


@app.get("/health/llm")
def health_llm():
    return llm.health()


@app.get("/api/template/costs.csv")
def template_costs():
    return Response(("\ufeff" + core.costs_template_csv()).encode("utf-8"), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": 'attachment; filename="costs_template.csv"'})


# ------------------------------------------------------------------ загрузка
@app.post("/api/upload")
def upload(reports: list[UploadFile] = File(default=[]), costs: UploadFile | None = File(default=None),
           consent: str = Form(""), tax_mode: str = Form("usn6"), tax_rate: str = Form(""),
           income_base: str = Form("revenue"), alloc: str = Form("revenue"),
           x_user_id: str = Header(default="")):
    if consent.lower() not in ("1", "true", "on", "yes"):
        raise HTTPException(400, "Нужно согласие на обработку данных.")
    uid = uid_of(x_user_id)
    limit("up:" + uid, 20, 3600)
    reports = [r for r in reports if r and r.filename]
    if len(reports) > MAX_FILES:
        raise HTTPException(400, f"Не больше {MAX_FILES} файлов отчётов за раз.")
    if not reports and not (costs and costs.filename):
        raise HTTPException(400, "Выберите хотя бы один файл.")
    W, errors, parsed, seen = [], [], [], set()
    for f in reports:
        data = read_upload(f)
        sha = hashlib.sha256(data).hexdigest()
        if sha in seen:
            W.append(("info", f"Файл «{f.filename}» загружен повторно, дубль пропущен."))
            continue
        seen.add(sha)
        try:
            parsed.append(core.parse_report(f.filename, data, W))
        except core.ParseError as e:
            errors.append(str(e))
        except Exception as e:  # noqa: BLE001
            errors.append(f"Файл «{f.filename}»: непредвиденная ошибка разбора ({type(e).__name__}).")
    new_costs = None
    if costs and costs.filename:
        try:
            new_costs = core.parse_costs(costs.filename, read_upload(costs), W)
        except core.ParseError as e:
            errors.append(str(e))
        except Exception as e:  # noqa: BLE001
            errors.append(f"Таблица себестоимости: непредвиденная ошибка разбора ({type(e).__name__}).")
    if reports and not parsed and new_costs is None:
        raise HTTPException(422, " ".join(errors) or "Не удалось разобрать файлы.")
    ops = core.combine_reports(parsed, W) if parsed else None
    st = clean_settings(tax_mode, tax_rate, income_base, alloc)
    old = load(uid)[2]
    st["parse_warnings"] = [list(w) for w in W] if ops is not None else old.get("parse_warnings", [])
    if ops is None and new_costs is not None:  # обновили только себестоимость
        st["parse_warnings"] = old.get("parse_warnings", []) + [list(w) for w in W]
    st["parse_warnings"] += [["error", e] for e in errors]
    with _lock:
        save(uid, ops, new_costs, st)
    return JSONResponse(build(uid))


@app.post("/api/settings")
def settings(body: dict, x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    key = key_for(uid, demo)
    old = load(key)[2]
    new = clean_settings(body.get("tax_mode"), body.get("tax_rate", ""), body.get("income_base"), body.get("alloc"), body.get("low_margin_pct"))
    old.update(new)
    if "tax_rate" in new and new["tax_rate"] is None:
        old["tax_rate"] = None
    save(key, None, None, old)
    return JSONResponse(build(key, bool(demo)))


@app.get("/api/analysis")
def analysis(x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    return JSONResponse(build(key_for(uid, demo), bool(demo)))


_demo_cache = {}


@app.post("/api/demo")
def demo(x_user_id: str = Header(default="")):
    uid = uid_of(x_user_id)
    limit("demo:" + uid, 30, 3600)
    if "v" not in _demo_cache:
        with tempfile.TemporaryDirectory() as td:
            synthetic.generate(td, articles=30, seed=11)
            W, parsed = [], []
            for name in ("wb_report_synthetic.xlsx", "ozon_accruals_synthetic.xlsx"):
                parsed.append(core.parse_report(name, (Path(td) / name).read_bytes(), W))
            c = core.parse_costs("costs.csv", (Path(td) / "costs_synthetic.csv").read_bytes(), W)
            _demo_cache["v"] = (core.combine_reports(parsed, W), c, W)
    ops, costs, W = _demo_cache["v"]
    key = "demo:" + uid
    save(key, ops, costs, {"tax_mode": "usn6", "parse_warnings": [list(w) for w in W]})
    return JSONResponse(build(key, True))


# ------------------------------------------------------------------ детали, сверка, экспорт
@app.get("/api/rows")
def rows(platform: str, article: str, x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    key = key_for(uid, demo)
    ops, costs, _, _ = load(key)
    if ops.empty:
        return []
    _, idmap = core.expand_costs(costs)
    art = article.strip().upper()
    m = (ops["article"].fillna("") == art) & (ops["platform"] == platform)
    if art.startswith("ID:"):
        m = (ops["ext_id"] == art[3:]) & (ops["platform"] == platform) & (ops["article"].fillna("") == "")
    else:
        ids = [e for (p, e), a in idmap.items() if p == platform and a == art]
        if ids:
            m |= (ops["ext_id"].isin(ids)) & (ops["platform"] == platform) & (ops["article"].fillna("") == "")
    out = ops[m].sort_values(["date", "src_row"]).head(1000)
    cols = ["date", "op_type", "order_id", "revenue", "commission", "acquiring", "logistics", "return_fees",
            "penalty", "storage", "ads", "other", "src_file", "src_row"]
    return JSONResponse(core.clean(out[cols].to_dict("records")))


@app.post("/api/payouts")
def payouts(body: dict, x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    key = key_for(uid, demo)
    plat, k = str(body.get("platform", "")), str(body.get("key", ""))
    if plat not in ("WB", "OZON") or not k or len(k) > 40:
        raise HTTPException(400, "Некорректные данные.")
    raw = body.get("amount")
    with closing(db()) as con:
        if raw in (None, ""):
            con.execute("DELETE FROM actuals WHERE user_id=? AND platform=? AND key=?", (key, plat, k))
        else:
            try:
                amt = float(str(raw).replace(" ", "").replace(",", "."))
            except ValueError:
                raise HTTPException(400, "Сумма должна быть числом.")
            con.execute("INSERT INTO actuals VALUES(?,?,?,?) ON CONFLICT(user_id, platform, key) DO UPDATE SET amount=excluded.amount",
                        (key, plat, k, amt))
        con.commit()
    return JSONResponse(build(key, bool(demo)))


@app.get("/api/export.csv")
def export_csv(x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    a = build(key_for(uid, demo), bool(demo))
    if a.get("empty"):
        raise HTTPException(404, "Нет данных.")
    cols = ["platform", "article", "name", "qty_sold", "qty_returned", "return_rate", "revenue", "commission", "acquiring",
            "logistics", "return_fees", "penalty", "storage", "ads", "other", "cogs", "profit_direct", "alloc", "profit_alloc",
            "margin", "tax_est", "profit_after_tax", "status"]
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(cols)
    for r in a["articles"]:
        w.writerow([str(r.get(c, "")).replace(".", ",") if isinstance(r.get(c), float) else r.get(c, "") for c in cols])
    return Response(("\ufeff" + buf.getvalue()).encode("utf-8"), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": 'attachment; filename="profit_by_article.csv"'})


@app.get("/api/report.pdf")
def report_pdf_ep(x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    limit("pdf:" + uid, 20, 3600)
    a = build(key_for(uid, demo), bool(demo))
    if a.get("empty"):
        raise HTTPException(404, "Нет данных.")
    try:
        data = report_pdf.build_pdf(a)
    except RuntimeError as e:
        raise HTTPException(500, str(e))
    return Response(data, media_type="application/pdf", headers={"Content-Disposition": 'attachment; filename="profit_report.pdf"'})


# ------------------------------------------------------------------ чат
@app.post("/api/chat")
def chat(body: dict, x_user_id: str = Header(default=""), demo: int = Query(0)):
    uid = uid_of(x_user_id)
    limit("chat:" + uid, 20, 600)
    if not llm.configured():
        raise HTTPException(503, "Чат-помощник не настроен на сервере (нет ключа GigaChat).")
    a = build(key_for(uid, demo), bool(demo))
    if a.get("empty"):
        raise HTTPException(400, "Сначала загрузите отчёт или запустите демо.")
    msg = str(body.get("message", "")).strip()[:1000]
    if not msg:
        raise HTTPException(400, "Пустой вопрос.")
    hist = [h for h in (body.get("history") or [])[-6:] if isinstance(h, dict) and h.get("role") in ("user", "assistant")]
    msgs = [{"role": "system", "content": llm.SYSTEM + "\n\nДАННЫЕ:\n" + llm.context(a)}]
    msgs += [{"role": h["role"], "content": str(h.get("content", ""))[:1000]} for h in hist]
    msgs.append({"role": "user", "content": msg})
    try:
        return {"answer": llm.ask(msgs)}
    except Exception as e:  # noqa: BLE001
        print("llm error:", type(e).__name__, str(e)[:200])
        raise HTTPException(502, "Не удалось получить ответ от GigaChat. Проверьте /health/llm.")


# ------------------------------------------------------------------ удаление данных
@app.delete("/api/data")
def delete_data(x_user_id: str = Header(default="")):
    if not UID_RE.match(x_user_id or ""):
        raise HTTPException(400, "Некорректный идентификатор.")
    with closing(db()) as con:
        wipe(con, x_user_id)
        con.commit()
    return {"ok": True}
