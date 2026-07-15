"""
SalesFlow — билдер площадки GGSel.

Источник: выгрузка GGSel (лист 'Chart data'), 1 строка = 1 продажа.
Сумма = 'Order Amount, USD'. Валюта = USD. Кол-во = sum(QTY).

Маппинг -> каталог по паре (база игры, тир издания):
- листинги 'Выбор издания' содержат издание в колонке 'Option name'
  ('… - Стандартное/Делюкс/Ультимейт издание'); эталон же кодирует издание
  суффиксом ('PRAGMATA USD' = стандарт, 'PRAGMATA: Deluxe Edition USD' = делюкс).
  Поэтому матчим по (base, tier), а не по сырому имени.
- база берётся из Option name, при неудаче — из Goods; тир — из Option name.
Несопоставленные (новые/не из каталога) помечаются пустым ID.
"""
from __future__ import annotations
import pandas as pd, re
from typing import Optional

CONST_PARTNER_GGSEL="Физическое лицо PLATI USD"
UNIFIED_COLUMNS=["Дата","ID","Наименование","Партнер","Количество","Цена","Валюта","Сумма"]
SUM_COL="Order Amount, USD"; QTY_COL="QTY"; GOODS_COL="Goods"; OPT_COL="Option name"

_TIERS=[("deluxe",r"делюкс|deluxe|dlx"),("ultimate",r"ультимейт|ultimate"),
        ("gold",r"\bgold\b|золот|голд"),("premium",r"премиум|premium"),
        ("complete",r"complete|комплит|полн\w*|definitive"),
        ("goty",r"goty|game of the year"),
        ("standard",r"стандарт\w*|standard|базов\w*|обычн\w*|\bstd\b")]
_STRIP=re.compile(r"\b(делюкс\w*|deluxe|dlx|ультимейт\w*|ultimate|gold|золот\w*|голд|"
                  r"премиум\w*|premium|complete|комплит\w*|полн\w*|definitive|goty|"
                  r"game of the year|стандарт\w*|standard|базов\w*|обычн\w*|std|"
                  r"издани\w*|edition|выбор|usd|eur|rub|rus|ru|смотреть описани\w*)\b",re.I)
_REG=re.compile(r"рф|снг|global|europe|весь мир|украин|без рф|\bес\b|\bтр\b|\bрб\b|\bкз\b|"
                r"индия|индонези|вьетнам|росси|страны",re.I)

def _norm(s)->str: return re.sub(r"\s+"," ",str(s)).strip().lower()
def _tier(s)->str:
    s=_norm(s)
    for t,pat in _TIERS:
        if re.search(pat,s): return t
    return "standard"
def _base(s)->str:
    s=_norm(s); s=re.sub(r"[™®©:]"," ",s)
    parts=[p.strip() for p in re.split(r"[|]",s)]
    parts=[p for p in parts if p and not re.fullmatch(r"(steam\s+)?(ключ|key)",p)]
    if len(parts)>1:
        parts=[p for p in parts if not _REG.search(p)] or parts
    s=max(parts,key=len) if parts else s
    s=_REG.sub(" ",s); s=_STRIP.sub(" ",s)
    return re.sub(r"[^\w]+"," ",s).strip()

def _find_sheet(path)->str:
    xls=pd.ExcelFile(path)
    for s in xls.sheet_names:
        cols=pd.read_excel(path,sheet_name=s,nrows=0,engine="calamine").columns
        if SUM_COL in cols and GOODS_COL in cols: return s
    return xls.sheet_names[0]

def load_ggsel_raw(path)->pd.DataFrame:
    df=pd.read_excel(path,sheet_name=_find_sheet(path),engine="calamine")
    df[SUM_COL]=pd.to_numeric(df[SUM_COL],errors="coerce")
    df[QTY_COL]=pd.to_numeric(df[QTY_COL],errors="coerce")
    return df

def build_ggsel_seed(etalon_path)->dict:
    et=pd.read_excel(etalon_path,sheet_name="ggsel",header=0,engine="calamine")
    et=et[pd.to_numeric(et["ID"],errors="coerce").notna()]
    seed={}
    for _,r in et.iterrows():
        for v in (r["Продукт"],r["Названия с площадки"]):
            if pd.notna(v): seed.setdefault((_base(v),_tier(v)),(int(r["ID"]),r["Продукт"]))
    return seed

def build_ggsel(raw_path, seed:dict,
                report_date:Optional[pd.Timestamp]=None)->pd.DataFrame:
    raw=load_ggsel_raw(raw_path)
    def look(row):
        opt=row[OPT_COL]; goods=row[GOODS_COL]; cands=[]
        if pd.notna(opt) and str(opt).strip():
            cands+=[(_base(opt),_tier(opt)),(_base(goods),_tier(opt))]
        cands.append((_base(goods),_tier(goods)))
        for k in cands:
            if k in seed: return seed[k]
        return None
    res=raw.apply(look,axis=1)
    raw["catID"]=[r[0] if r else pd.NA for r in res]
    raw["product"]=[r[1] if r else None for r in res]
    raw["_unmapped"]=raw["catID"].isna()
    raw["Наименование"]=raw["product"].where(~raw["_unmapped"],raw[GOODS_COL])
    g=raw.groupby(["catID","Наименование"],dropna=False,as_index=False).agg(
        Количество=(QTY_COL,"sum"),Сумма=(SUM_COL,"sum"))
    g["Цена"]=(g["Сумма"]/g["Количество"]).where(g["Количество"]!=0,0).round(6)
    g["Сумма"]=g["Сумма"].round(2)
    out=pd.DataFrame({
        "Дата":pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID":pd.array(g["catID"],dtype="Int64"),
        "Наименование":g["Наименование"].astype("string"),
        "Партнер":CONST_PARTNER_GGSEL,
        "Количество":g["Количество"].astype("Int64"),
        "Цена":g["Цена"],"Валюта":"USD","Сумма":g["Сумма"],
    })[UNIFIED_COLUMNS]
    return out.sort_values(["ID"]).reset_index(drop=True)
