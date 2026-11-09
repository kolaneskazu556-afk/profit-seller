# -*- coding: utf-8 -*-
"""PDF-отчёт (reportlab). Нужен шрифт DejaVu (пакет fonts-dejavu-core на сервере)."""
import io
import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from xml.sax.saxutils import escape

FONT_DIRS = ["/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/dejavu", "C:/Windows/Fonts"]


def _fonts():
    for d in FONT_DIRS:
        r, b = os.path.join(d, "DejaVuSans.ttf"), os.path.join(d, "DejaVuSans-Bold.ttf")
        if os.path.exists(r) and os.path.exists(b):
            pdfmetrics.registerFont(TTFont("DV", r))
            pdfmetrics.registerFont(TTFont("DV-B", b))
            return
    raise RuntimeError("Не найден шрифт DejaVuSans (установите fonts-dejavu-core).")


def _n(x):
    return "—" if x is None else f"{x:,.0f}".replace(",", " ")


def build_pdf(a: dict) -> bytes:
    _fonts()
    st = lambda n, **k: ParagraphStyle(n, fontName=k.pop("f", "DV"), **k)  # noqa: E731
    h1, h2 = st("h1", f="DV-B", fontSize=16, leading=20, spaceAfter=6), st("h2", f="DV-B", fontSize=12, leading=15, spaceBefore=10, spaceAfter=4)
    p, small = st("p", fontSize=9, leading=12), st("s", fontSize=7.5, leading=10, textColor=colors.grey)
    P = lambda t, s=p: Paragraph(escape(str(t)), s)  # noqa: E731
    t, m = a["totals"], a["meta"]
    el = [P("Отчёт о прибыли по артикулам", h1),
          P(f"Период: {m['period_from']} — {m['period_to']}. Площадки: {', '.join(m['platforms'])}. "
            + ("ДЕМО: данные вымышленные." if m.get("demo") else ""), p), Spacer(1, 4)]
    rows = [["Показатель", "₽"], ["Выручка (с учётом возвратов)", _n(t["revenue"])]]
    for b in ["commission", "acquiring", "logistics", "return_fees", "penalty", "storage", "ads", "other"]:
        rows.append([a["labels"][b], "−" + _n(t[b])])
    rows += [["Себестоимость проданного", "−" + _n(t["cogs"])], ["Прибыль до налога", _n(t["profit_before_tax"])],
             [f"Налог (ориентир., {a['tax']['name']}, {a['tax']['rate']}%)", "−" + _n(t["tax"])],
             ["Прибыль после налога (ориентир.)", _n(t["profit_after_tax"])]]
    tb = Table(rows, colWidths=[110 * mm, 40 * mm])
    tb.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "DV", 9), ("FONT", (0, 0), (-1, 0), "DV-B", 9),
                            ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ffe6b0"))]))
    el += [tb, P("Выводы", h2)] + [P("• " + x) for x in a["insights"]]
    loss = [x for x in a["articles"] if x["status"] == "loss"][:20]
    if loss:
        el.append(P("Убыточные артикулы", h2))
        for x in loss:
            el.append(P("• " + x["explain"]))
    top = a["articles"][:25]
    rows = [["Пл.", "Артикул", "Прод.", "Возвр.%", "Выручка", "Прибыль", "Маржа"]]
    for x in top:
        rows.append([x["platform"], x["article"], f"{x['qty_sold']:.0f}", f"{x['return_rate']:.0f}", _n(x["revenue"]),
                     _n(x["profit_direct"]), "—" if x["margin"] is None else f"{x['margin'] * 100:.0f}%"])
    tb = Table(rows, repeatRows=1)
    tb.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "DV", 8), ("FONT", (0, 0), (-1, 0), "DV-B", 8),
                            ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.lightgrey)]))
    el += [P("Худшие 25 артикулов по прямой прибыли", h2), tb, P("Как считалось и ограничения", h2)]
    el += [P("• " + x, small) for x in a["methodology"] + a["tax"]["notes"]]
    if a["warnings"]:
        el.append(P("Предупреждения по данным", h2))
        el += [P("• " + w["text"], small) for w in a["warnings"]]
    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=15 * mm, bottomMargin=15 * mm,
                      title="Отчёт о прибыли").build(el)
    return buf.getvalue()
