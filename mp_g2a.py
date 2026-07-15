"""
SalesFlow — билдер площадки G2A.

Источник: выгрузка кабинета G2A (лист 'report_statement_*').
Продажи = строки Type == 'Product' (включая редкие Payment Reversal с Qty<0,
которые нетто-вычитаются). Сумма = 'Amount EUR (Approx,)' — готовая EUR-конвертация
G2A (валют продажи много, эталон берёт именно её). Кол-во = sum(Qty).

Маппинг листинга -> каталог по 'Name' (нормализованному), засев из эталона.
"""
from __future__ import annotations
import pandas as pd, re
from typing import Optional

CONST_PARTNER_G2A="Физическое лицо G2A"
UNIFIED_COLUMNS=["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]
TYPE_COL="Type"; SALE_TYPE="Product"; EUR_COL="Amount EUR (Approx,)"
QTY_COL="Qty"; NAME_COL="Name"

def _norm(s)->str: return re.sub(r"\s+"," ",str(s)).strip().lower()

def _find_sheet(path)->str:
    xls=pd.ExcelFile(path)
    for s in xls.sheet_names:
        if s.lower().startswith("report_statement"): return s
    best,best_rows=None,-1
    for s in xls.sheet_names:
        cols=pd.read_excel(path,sheet_name=s,nrows=0,engine="calamine").columns
        if TYPE_COL in cols and EUR_COL in cols:
            n=pd.read_excel(path,sheet_name=s,usecols=[0],engine="calamine").shape[0]
            if n>best_rows: best,best_rows=s,n
    return best or xls.sheet_names[0]

def load_g2a_raw(path)->pd.DataFrame:
    df=pd.read_excel(path,sheet_name=_find_sheet(path),engine="calamine")
    df[EUR_COL]=pd.to_numeric(df[EUR_COL],errors="coerce")
    df[QTY_COL]=pd.to_numeric(df[QTY_COL],errors="coerce")
    return df

def build_g2a_mapping_from_etalon(etalon_path)->pd.DataFrame:
    et=pd.read_excel(etalon_path,sheet_name="G2A",header=0,engine="calamine")
    rb=et.iloc[:,10:14].copy(); rb.columns=["catID","listing","k","s"]
    rb=rb[pd.to_numeric(rb["catID"],errors="coerce").notna()]
    lb=et.iloc[:,0:6].copy(); lb.columns=["catID","product","k","p","c","s"]
    lb=lb[pd.to_numeric(lb["catID"],errors="coerce").notna()]
    mp=(rb[["listing","catID"]].merge(lb[["catID","product"]],on="catID",how="left")
        .drop_duplicates("listing"))
    mp["catID"]=mp["catID"].astype("Int64")
    return mp[["listing","catID","product"]]

def build_g2a(raw_path, mapping:pd.DataFrame,
              report_date:Optional[pd.Timestamp]=None)->pd.DataFrame:
    raw=load_g2a_raw(raw_path)
    s=raw[raw[TYPE_COL]==SALE_TYPE].copy()
    m=mapping.dropna(subset=["listing"]).copy(); m["_k"]=m["listing"].map(_norm)
    nm=m.drop_duplicates("_k").set_index("_k")
    s["_k"]=s[NAME_COL].map(_norm)
    s["catID"]=s["_k"].map(nm["catID"]); s["product"]=s["_k"].map(nm["product"])
    s["_unmapped"]=s["catID"].isna()
    s["Наименование"]=s["product"].where(~s["_unmapped"],s[NAME_COL])
    g=s.groupby(["catID","Наименование"],dropna=False,as_index=False).agg(
        Количество=(QTY_COL,"sum"),Сумма=(EUR_COL,"sum"))
    g["Цена"]=(g["Сумма"]/g["Количество"]).where(g["Количество"]!=0,0).round(6)
    g["Сумма"]=g["Сумма"].round(2)
    out=pd.DataFrame({
        "Дата":pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID":g["catID"].astype("Int64"),
        "Наименование":g["Наименование"].astype("string"),
        "Партнер":CONST_PARTNER_G2A,
        "Количество":g["Количество"].astype("Int64"),
        "Цена":g["Цена"],"Валюта":"EUR","Сумма":g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["ID"]).reset_index(drop=True)
