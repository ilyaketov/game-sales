"""
SalesFlow — билдер площадки Plati (+ суб-магазины WMZ/WMT/WMR).

Источник: выгрузка Plati (лист 'выгрузка'), 1 строка = 1 ключ.
Фильтры: площадка == 'Plati'; исключается товар 'АВТОПОПОЛНЕНИЕ STEAM'
(автопополнение кошелька — не продажа ключа; именно это расходилось с эталоном).
Сумма = 'зачислено'. Валюта: WMR -> RUB, WMT/WMZ -> USD.

Маппинг листинга -> каталог по 'Name' (нормализованному), засев из NUKEN-листов
эталона. Бандлы 'Выбор издания' маппятся не 1:1 — помечаются для ручной разбивки.
"""
from __future__ import annotations
import pandas as pd, re
from typing import Optional

UNIFIED_COLUMNS=["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]
EXCLUDE_NAME="автопополнение steam"
CUR_MAP={"WMR":"RUB","WMT":"USD","WMZ":"USD"}
PARTNER_BY_CUR={"RUB":"Физическое лицо PLATI rub","USD":"Физическое лицо PLATI USD"}

def _norm(s)->str: return re.sub(r"\s+"," ",str(s)).strip().lower()

def load_plati_raw(path)->pd.DataFrame:
    xls=pd.ExcelFile(path)
    sheet="выгрузка" if "выгрузка" in xls.sheet_names else max(
        xls.sheet_names,key=lambda s: pd.read_excel(path,sheet_name=s,nrows=1,engine="calamine").shape[1])
    df=pd.read_excel(path,sheet_name=sheet,engine="calamine")
    df["зачислено"]=pd.to_numeric(df["зачислено"],errors="coerce")
    return df

def build_plati_mapping_from_etalon(etalon_path)->pd.DataFrame:
    maps=[]
    for sh in ["PLATIWMZ NUKEN","PLATIWMT NUKEN","PLATIRUB NUKEN"]:
        try: et=pd.read_excel(etalon_path,sheet_name=sh,header=0,engine="calamine")
        except Exception: continue
        sub=et[["ID","Продукт","Названия с площадки"]].copy()
        sub=sub[pd.to_numeric(sub["ID"],errors="coerce").notna()]
        sub.columns=["catID","product","listing"]; maps.append(sub)
    mp=pd.concat(maps).dropna(subset=["listing"]).drop_duplicates("listing")
    mp["catID"]=mp["catID"].astype("Int64")
    return mp[["listing","catID","product"]]

def build_plati(raw_path, mapping:pd.DataFrame,
                report_date:Optional[pd.Timestamp]=None)->pd.DataFrame:
    raw=load_plati_raw(raw_path)
    d=raw[raw["площадка"]=="Plati"].copy()
    d=d[d["Name"].map(_norm)!=EXCLUDE_NAME]
    d["Валюта"]=d["валюта"].map(CUR_MAP)
    nm=mapping.dropna(subset=["listing"]).copy(); nm["_k"]=nm["listing"].map(_norm)
    nm=nm.drop_duplicates("_k").set_index("_k")
    d["_k"]=d["Name"].map(_norm)
    d["catID"]=d["_k"].map(nm["catID"]); d["product"]=d["_k"].map(nm["product"])
    d["_unmapped"]=d["catID"].isna()
    d["Наименование"]=d["product"].where(~d["_unmapped"],d["Name"])
    g=d.groupby(["catID","Наименование","Валюта"],dropna=False,as_index=False).agg(
        Количество=("зачислено","size"),Сумма=("зачислено","sum"))
    g["Цена"]=(g["Сумма"]/g["Количество"]).where(g["Количество"]!=0,0).round(6)
    g["Сумма"]=g["Сумма"].round(2)
    out=pd.DataFrame({
        "Дата":pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID":pd.array(g["catID"],dtype="Int64"),
        "Наименование":g["Наименование"].astype("string"),
        "Партнер":g["Валюта"].map(PARTNER_BY_CUR).astype("string"),
        "Количество":g["Количество"].astype("Int64"),
        "Цена":g["Цена"],"Валюта":g["Валюта"].astype("string"),"Сумма":g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["Валюта","ID"]).reset_index(drop=True)
