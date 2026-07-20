"""
SalesFlow — билдер площадки Eneba.

Источник: выгрузка кабинета Eneba (лист 'report*' с колонками Тип/Оплаченная сумма
ИЛИ их англ. эквивалентами Type/Paid amount — с июня 2026 кабинет отдаёт англ.
локаль). 1 строка = 1 проданный ключ.

Продажи: Тип == 'Продажа' / Type == 'Sale'. Сумма = 'Оплаченная сумма' /
'Paid amount' (EUR). Количество ключей = число строк. Строки прошлого месяца
(рус. метка валюты 'возврат <месяц>') исключаются, если такая метка присутствует
(в англ. выгрузке её нет — фильтр становится no-op).

Маппинг листинга -> каталог по 'name', таблица персистентная, засев из эталона.
"""
from __future__ import annotations
import pandas as pd
import re
from typing import Optional


def _norm(s) -> str:
    """Нормализация имени листинга для маппинга: trim, схлоп пробелов, lower."""
    return re.sub(r"\s+", " ", str(s)).strip().lower()

CONST_PARTNER_ENEBA = "Физическое лицо ENEBA"
UNIFIED_COLUMNS = ["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]

# Двуязычные псевдонимы колонок/значений (RU — старый формат, EN — с июня 2026).
TYPE_ALIASES     = ["Тип", "Type"]
SALE_VALUES      = {"продажа", "sale"}
PAY_ALIASES      = ["Оплаченная сумма", "Paid amount"]
PAY_CUR_ALIASES  = ["Оплаченная сумма - валюта", "Paid amount - currency"]
NAME_COL = "name"
SKIP_CUR_PREFIX = "возврат"   # рус. метка возврата прошлого месяца (напр. 'возврат мая')


def _resolve(cols, aliases):
    """Первый псевдоним, реально присутствующий в cols (case/space-insensitive)."""
    norm = {str(c).strip().lower(): c for c in cols}
    for a in aliases:
        if a.strip().lower() in norm:
            return norm[a.strip().lower()]
    return None


def _find_tx_sheet(path) -> str:
    xls = pd.ExcelFile(path, engine="calamine")
    best, best_rows = None, -1
    for s in xls.sheet_names:
        cols = pd.read_excel(path, sheet_name=s, nrows=0, engine="calamine").columns
        if _resolve(cols, TYPE_ALIASES) and _resolve(cols, PAY_ALIASES):
            n = pd.read_excel(path, sheet_name=s, usecols=[0], engine="calamine").shape[0]
            if n > best_rows:
                best, best_rows = s, n
    return best or xls.sheet_names[0]


def load_eneba_raw(path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=_find_tx_sheet(path), engine="calamine")
    pay = _resolve(df.columns, PAY_ALIASES)
    if pay is not None:
        df[pay] = pd.to_numeric(df[pay], errors="coerce")
    return df


def build_eneba_mapping_from_etalon(etalon_path) -> pd.DataFrame:
    et = pd.read_excel(etalon_path, sheet_name="ENEBA", header=0, engine="calamine")
    rb = et.iloc[:, 10:14].copy(); rb.columns = ["catID","listing","k","s"]
    rb = rb[pd.to_numeric(rb["catID"], errors="coerce").notna()]
    lb = et.iloc[:, 0:6].copy(); lb.columns = ["catID","product","k","p","c","s"]
    lb = lb[pd.to_numeric(lb["catID"], errors="coerce").notna()]
    mp = (rb[["listing","catID"]].merge(lb[["catID","product"]], on="catID", how="left")
          .drop_duplicates("listing"))
    mp["catID"] = mp["catID"].astype("Int64")
    return mp[["listing","catID","product"]]


def build_eneba(raw_path, mapping: pd.DataFrame,
                report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    raw = load_eneba_raw(raw_path)
    type_col = _resolve(raw.columns, TYPE_ALIASES)
    pay_col  = _resolve(raw.columns, PAY_ALIASES)
    pay_cur  = _resolve(raw.columns, PAY_CUR_ALIASES)
    if type_col is None or pay_col is None:
        raise ValueError("Eneba: не найдены колонки Тип/Type + Оплаченная сумма/Paid amount")

    is_sale = raw[type_col].astype(str).str.strip().str.lower().isin(SALE_VALUES)
    s = raw[is_sale].copy()
    # исключаем возвраты прошлого месяца, если рус. метка присутствует в колонке валюты
    if pay_cur is not None:
        cur = s[pay_cur].astype(str).str.strip().str.lower()
        s = s[~cur.str.startswith(SKIP_CUR_PREFIX)].copy()

    m = mapping.dropna(subset=["listing"]).copy()
    m["_k"] = m["listing"].map(_norm)
    nm = m.drop_duplicates("_k").set_index("_k")
    s["_k"] = s[NAME_COL].map(_norm)
    s["catID"] = s["_k"].map(nm["catID"]); s["product"] = s["_k"].map(nm["product"])
    s["_unmapped"] = s["catID"].isna()
    s["Наименование"] = s["product"].where(~s["_unmapped"], s[NAME_COL])
    g = s.groupby(["catID","Наименование"], dropna=False, as_index=False).agg(
        Количество=(pay_col, "size"), Сумма=(pay_col, "sum"))
    g["Цена"] = (g["Сумма"] / g["Количество"]).where(g["Количество"] != 0, 0).round(6)
    g["Сумма"] = g["Сумма"].round(2)
    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": g["catID"].astype("Int64"),
        "Наименование": g["Наименование"].astype("string"),
        "Партнер": CONST_PARTNER_ENEBA,
        "Количество": g["Количество"].astype("Int64"),
        "Цена": g["Цена"], "Валюта": "EUR", "Сумма": g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["ID"]).reset_index(drop=True)
