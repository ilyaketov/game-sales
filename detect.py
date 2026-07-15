"""Распознавание роли загруженного файла по сигнатуре колонок.

Headless и тестируемо (не зависит от Streamlit). Используется единым окном
загрузки в app.py: кабинеты распознаются первыми (специфичнее), затем биллинг,
затем опциональные Остаток_нач/События.
"""
from __future__ import annotations

import pandas as pd
import config

KIND_LABEL = {
    "r1": "R1 — Универсальный отчёт", "r2": "R2 — shipped", "genba": "genbaFile",
    "Eneba": "Выгрузка Eneba", "Kinguin": "Выгрузка Kinguin",
    "Driffle": "Выгрузка Driffle", "G2A": "Выгрузка G2A",
    "Plati": "Выгрузка Plati", "GGSel": "Выгрузка GGSel",
    "carry": "Остаток_нач (перенос)", "events": "События (перемещения)",
}
CORE = ["r1", "r2", "genba", "Eneba", "Kinguin", "Driffle", "G2A", "Plati", "GGSel"]


def _norm(c) -> str:
    return str(c).strip().lower()


_SIG = {k: {_norm(v) for v in getattr(config, n).values()}
        for k, n in (("r1", "COLS_R1"), ("r2", "COLS_R2"), ("genba", "COLS_GENBA"))}


def file_columns(path: str) -> set:
    """Объединение нормализованных заголовков всех листов файла."""
    cols: set = set()
    if str(path).lower().endswith(".csv"):
        try:
            cols |= {_norm(c) for c in pd.read_csv(path, nrows=0).columns}
        except Exception:
            pass
        return cols
    try:
        xl = pd.ExcelFile(path, engine="calamine")
    except Exception:
        return cols
    for sh in xl.sheet_names:
        try:
            h = pd.read_excel(xl, sheet_name=sh, nrows=0, engine="calamine")
            cols |= {_norm(c) for c in h.columns}
        except Exception:
            pass
    return cols


def detect_kind(path: str) -> str | None:
    """Роль файла или None. Кабинеты — первыми (их колонки уникальны)."""
    c = file_columns(path)
    if not c:
        return None
    if "productid" in c and "price" in c:
        return "Kinguin"
    if any("order amount" in x and "usd" in x for x in c) and \
       (any("option name" in x for x in c) or any("goods" in x for x in c)):
        return "GGSel"
    if any("selling price" in x for x in c) and any("product title" in x for x in c):
        return "Driffle"
    if any(x.startswith("amount eur") for x in c):
        return "G2A"
    if any("оплаченная сумма" in x for x in c) and "тип" in c:
        return "Eneba"
    if any("зачислено" in x for x in c):
        return "Plati"
    if len(c & _SIG["r2"]) >= 10:
        return "r2"
    if len(c & _SIG["r1"]) >= 8:
        return "r1"
    if len(c & _SIG["genba"]) >= 3:
        return "genba"
    if "события" in c and "id" in c and len(c) <= 6:
        return "events"
    if ("остаток_нач" in c or "остаток_конец" in c) and "id" in c and len(c) <= 6:
        return "carry"
    return None
