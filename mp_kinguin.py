"""
SalesFlow — билдер площадки Kinguin.

Источник: сырая выгрузка кабинета Kinguin (лист reservations_*),
где 1 строка = 1 доставленный ключ.

Выход: стандартный лист SalesFlow
    [Дата, ID, Наименование, Партнер, Количество, Цена, Валюта, Сумма]

Маппинг листинга Kinguin -> внутренний каталог идёт по PRODUCTID (стабильный
хеш Kinguin), с фолбэком на NAME. Таблица маппинга персистентная, засевается
из эталона один раз и доращивается новыми листингами.
"""
from __future__ import annotations
import pandas as pd
from pathlib import Path
from typing import Optional

CONST_PARTNER_KINGUIN = "Физическое лицо Kinguin"
UNIFIED_COLUMNS = ["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]


def _find_reservations_sheet(path) -> str:
    xls = pd.ExcelFile(path)
    for s in xls.sheet_names:
        if s.lower().startswith("reservation"):
            return s
    # фолбэк: лист с колонкой PRICE и максимумом строк
    best, best_rows = None, -1
    for s in xls.sheet_names:
        cols = pd.read_excel(path, sheet_name=s, nrows=0, engine="calamine").columns
        if "PRICE" in cols:
            n = pd.read_excel(path, sheet_name=s, usecols=[0], engine="calamine").shape[0]
            if n > best_rows:
                best, best_rows = s, n
    return best or xls.sheet_names[0]


def load_kinguin_raw(path) -> pd.DataFrame:
    sheet = _find_reservations_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet, engine="calamine")
    df["PRICE"] = pd.to_numeric(df["PRICE"], errors="coerce")
    return df


def build_kinguin_mapping_from_etalon(etalon_path, raw_path) -> pd.DataFrame:
    """Сеем таблицу маппинга PRODUCTID/NAME -> внутренний (ID, Продукт) из эталона.

    Эталон даёт пару (листинг 'Названия с площадки' -> внутренний ID, Продукт).
    Сырьё даёт пару (NAME -> PRODUCTID). Джойн по имени листинга связывает
    стабильный PRODUCTID с внутренним ID.
    """
    et = pd.read_excel(etalon_path, sheet_name="Kinguin", header=0, engine="calamine")
    rb = et.iloc[:, 10:14].copy(); rb.columns = ["catID","listing","k","s"]
    rb = rb[pd.to_numeric(rb["catID"], errors="coerce").notna()]
    lb = et.iloc[:, 0:6].copy(); lb.columns = ["catID","product","k","p","c","s"]
    lb = lb[pd.to_numeric(lb["catID"], errors="coerce").notna()]
    name_to_cat = (rb[["listing","catID"]]
                   .merge(lb[["catID","product"]], on="catID", how="left")
                   .drop_duplicates("listing"))
    name_to_cat["catID"] = name_to_cat["catID"].astype("Int64")

    raw = load_kinguin_raw(raw_path)
    name_to_pid = raw[["NAME","PRODUCTID"]].drop_duplicates("NAME")

    mp = name_to_pid.merge(name_to_cat, left_on="NAME", right_on="listing", how="outer")
    mp = mp.rename(columns={"NAME":"listing_raw"})
    mp["listing"] = mp["listing"].fillna(mp["listing_raw"])
    return mp[["PRODUCTID","listing","catID","product"]]


def build_kinguin(raw_path, mapping: pd.DataFrame,
                  report_date: Optional[pd.Timestamp] = None,
                  status_filter: str = "DELIVERED") -> pd.DataFrame:
    raw = load_kinguin_raw(raw_path)
    if status_filter:
        raw = raw[raw["STATUS"] == status_filter]

    # маппинг: сначала по PRODUCTID, фолбэк по NAME
    pid_map = mapping.dropna(subset=["PRODUCTID"]).drop_duplicates("PRODUCTID").set_index("PRODUCTID")
    name_map = mapping.dropna(subset=["listing"]).drop_duplicates("listing").set_index("listing")

    raw = raw.copy()
    raw["catID"] = raw["PRODUCTID"].map(pid_map["catID"])
    raw["product"] = raw["PRODUCTID"].map(pid_map["product"])
    miss = raw["catID"].isna()
    raw.loc[miss, "catID"] = raw.loc[miss, "NAME"].map(name_map["catID"])
    raw.loc[miss, "product"] = raw.loc[miss, "NAME"].map(name_map["product"])

    # для несопоставленных оставляем листинговое имя, ID пустой, статус-флаг
    raw["_unmapped"] = raw["catID"].isna()
    raw["Наименование"] = raw["product"].where(~raw["_unmapped"], raw["NAME"])

    g = raw.groupby(["catID","Наименование","CURRENCY"], dropna=False, as_index=False).agg(
        Количество=("PRICE","size"), Сумма=("PRICE","sum"))
    g["Цена"] = (g["Сумма"]/g["Количество"]).round(6)
    g["Сумма"] = g["Сумма"].round(2)

    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": g["catID"].astype("Int64"),
        "Наименование": g["Наименование"].astype("string"),
        "Партнер": CONST_PARTNER_KINGUIN,
        "Количество": g["Количество"].astype("Int64"),
        "Цена": g["Цена"],
        "Валюта": g["CURRENCY"].astype("string"),
        "Сумма": g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["ID"]).reset_index(drop=True)
