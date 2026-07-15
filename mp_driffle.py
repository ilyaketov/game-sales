"""
SalesFlow — билдер площадки Driffle.

Источник: выгрузка кабинета Driffle (лист transaction_report_*),
1 строка = 1 транзакция. Продажи Type='Sale', возвраты Type='Refund'.

Нетто-продажи = Sale минус Refund. Сумма берётся из 'Selling Price (EUR)'
(валовая цена продажи — совпадает с эталоном). Валюта = EUR.

Маппинг листинга -> каталог по 'Product Title' (стабильного ID у Driffle нет),
таблица персистентная, засевается из эталона.
"""
from __future__ import annotations
import pandas as pd
from typing import Optional

CONST_PARTNER_DRIFFLE = "Физическое лицо DRIFFL"
UNIFIED_COLUMNS = ["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]
SUM_COL = "Selling Price (EUR)"
TITLE_COL = "Product Title"


def _find_tx_sheet(path) -> str:
    xls = pd.ExcelFile(path)
    for s in xls.sheet_names:
        if s.lower().startswith("transaction_report"):
            return s
    best, best_rows = None, -1
    for s in xls.sheet_names:
        cols = pd.read_excel(path, sheet_name=s, nrows=0, engine="calamine").columns
        if TITLE_COL in cols and SUM_COL in cols:
            n = pd.read_excel(path, sheet_name=s, usecols=[0], engine="calamine").shape[0]
            if n > best_rows:
                best, best_rows = s, n
    return best or xls.sheet_names[0]


def load_driffle_raw(path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=_find_tx_sheet(path), engine="calamine")
    df[SUM_COL] = pd.to_numeric(df[SUM_COL], errors="coerce")
    return df


def build_driffle_mapping_from_etalon(etalon_path) -> pd.DataFrame:
    """Сеем маппинг 'Product Title' -> (catID, Продукт) из листа Driffle эталона."""
    et = pd.read_excel(etalon_path, sheet_name="Driffle", header=0, engine="calamine")
    rb = et.iloc[:, 10:14].copy(); rb.columns = ["catID","listing","k","s"]
    rb = rb[pd.to_numeric(rb["catID"], errors="coerce").notna()]
    lb = et.iloc[:, 0:6].copy(); lb.columns = ["catID","product","k","p","c","s"]
    lb = lb[pd.to_numeric(lb["catID"], errors="coerce").notna()]
    mp = (rb[["listing","catID"]]
          .merge(lb[["catID","product"]], on="catID", how="left")
          .drop_duplicates("listing"))
    mp["catID"] = mp["catID"].astype("Int64")
    return mp[["listing","catID","product"]]


def build_driffle(raw_path, mapping: pd.DataFrame,
                  report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    raw = load_driffle_raw(raw_path)
    sr = raw[raw["Type"].isin(["Sale","Refund"])].copy()
    sr["_qty"] = sr["Type"].map({"Sale": 1, "Refund": -1})

    name_map = mapping.dropna(subset=["listing"]).drop_duplicates("listing").set_index("listing")
    sr["catID"]   = sr[TITLE_COL].map(name_map["catID"])
    sr["product"] = sr[TITLE_COL].map(name_map["product"])
    sr["_unmapped"] = sr["catID"].isna()
    sr["Наименование"] = sr["product"].where(~sr["_unmapped"], sr[TITLE_COL])

    g = sr.groupby(["catID","Наименование"], dropna=False, as_index=False).agg(
        Количество=("_qty","sum"), Сумма=(SUM_COL,"sum"))
    g["Цена"] = (g["Сумма"]/g["Количество"]).where(g["Количество"]!=0, 0).round(6)
    g["Сумма"] = g["Сумма"].round(2)

    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": g["catID"].astype("Int64"),
        "Наименование": g["Наименование"].astype("string"),
        "Партнер": CONST_PARTNER_DRIFFLE,
        "Количество": g["Количество"].astype("Int64"),
        "Цена": g["Цена"],
        "Валюта": "EUR",
        "Сумма": g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["ID"]).reset_index(drop=True)
