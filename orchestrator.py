"""
SalesFlow unified — оркестратор единого прохода.

Вход (один месяц):
  - Universal Report R1  +  R2 shipped  +  genbaFile   -> KeyFlow -> Загружено по pid
  - 6 выгрузок витрин  (Kinguin/Driffle/Eneba/G2A/GGSel/Plati)  -> спрос + деньги
  - Каталог            (ID -> base, Регион, Название)
  - Остаток_нач        (перенос с прошлого месяца; для seed-месяца — из эталона)

Проход:
  1) zagruzheno()      KeyFlow.Pipeline.aggregate(зона) -> qty по pid  == Загружено
  2) catalog_map()     ID -> (base, Регион, Название)
  3) build_stock()     Остаток = Остаток_нач + Загружено  (+ Регион, base)
  4) sales_by_channel() 6 парсеров -> унифицированный лист + спрос по base
  5) allocate_region() для регион-агностичных каналов спрос -> региональные ID
  6) movement()        нач + загружено − продано = конец  (вычисляемое «Движение ключей»)
  7) review()          позиции с флагами allocator (~0,4%)

Контракт листа всех каналов (закуп и продажи) единый:
  [Дата, ID, Наименование, Партнер, Количество, Цена, Валюта, Сумма].

ВАЖНО (гейт валидации): точное определение «Загружено» — какие зоны KeyFlow
суммируются по pid (только закуп, или закуп + перемещение) — проверяется
функцией validate_zagruzheno() против колонки «Загружено» эталонного
«Движения ключей». До прохождения гейта оркестратор НЕ считается верным.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pandas as pd

import allocator
import engine_sales as sales          # build_b2b / build_chinaplay / build_gb, load_catalog
from engine_keyflow import Pipeline
from config import PLOSHADKA_MAP
import mp_kinguin, mp_driffle, mp_eneba, mp_g2a, mp_ggsel, mp_plati

# РАЗНЕСЕНИЕ — УНИВЕРСАЛЬНО ДЛЯ ВСЕХ КАНАЛОВ.
# ПОДТВЕРЖДЕНО (май, 4 канала): PRODUCTID (Kinguin) / название (Driffle) /
# Option name (GGSel) резолвят ИГРУ и ИЗДАНИЕ, но НЕ регион — региональный ключ
# выбирается при выдаче со склада. Поэтому per-ID региональный сплит для
# «Движения ключей» даёт allocator на ВСЕХ каналах:
#   Eneba 99,6% · Kinguin 99,5% · Driffle 97,4% · G2A 93,5% (Σ везде точно).
# Билдер — источник ДЕНЕГ и ИТОГА по листингу; allocator — источник СПЛИТА по ID.
# Прежнее деление на «регион зашит / агностичный» снято как неверное.
# Общие реестры: несколько каналов делят ОДНО «Движение ключей».
# ПОДТВЕРЖДЕНО (май): Plati и GGSel — один сток, «Загружено» из KeyFlow-зоны
# «Plati», продажи разбиты по подканалам (ggsel + WMZ/WMT/RUB). GGSel НЕ имеет
# своей закуп-зоны. Тождество общего реестра закрылось 632/632, Σ|Δ|=0.
COMBINED_LEDGERS = {"Plati+GGSel": {"zone": "Plati", "channels": ["Plati", "GGSel"]}}

ALLOCATOR_CHANNELS = ["Eneba", "Kinguin", "Driffle", "G2A", "Plati", "GGSel"]

# Канал продаж (витрина) -> закуп-зона KeyFlow, формирующая её «Загружено».
# ПОДТВЕРЖДЕНО на Eneba (май): qty по pid из aggregate("Eneba") == эталон
# ЗАГРУЖЕНО 87/87, Σ|Δ|=0. «Загружено» в «Движении ключей» — КАНАЛЬНОЕ
# (закуп-зона своего канала), НЕ сумма всех зон (сумма всех = весь закуп).
CHANNEL_TO_ZONE = {
    "Eneba": "Eneba", "Kinguin": "Kinguin", "Driffle": "Driffle",
    "G2A": "G2A", "GGSel": "ggsel", "Plati": "Plati",
    "B2B": "B2B", "Chinaplay": "ChinaPlay", "GamersBase": "GamersBase",
}

REGION_RE = re.compile(
    r"\s+(RUB|RU|EUR|PLN|GBP|USD|KRW|UAH|KZT|TRY|TL|NOK|SEK|DKK|"
    r"BRL|ARS|JPY|CNY|INR|MXN|AUD|CAD|CHF)\s*$", re.I)


def _base_name(name: str) -> str:
    """Имя игры/издания без регион-суффикса (для группировки спроса)."""
    return REGION_RE.sub("", str(name)).strip()


# ---------------------------------------------------------------------------
# Стадия 1 — Загружено по каталоговому ID (из биллинга через KeyFlow)
# ---------------------------------------------------------------------------
def load_pipeline(r1: str, r2: str, genba: str) -> Pipeline:
    """Загружает биллинг один раз (R2/genba крупные ~40МБ, ~50с)."""
    return Pipeline(r1, r2, genba)


def zagruzheno(pipe: Pipeline, channel: str) -> pd.DataFrame:
    """«Загружено» по pid для ОДНОГО канала = qty закуп-зоны этого канала.

    Возврат: [ID, Загружено, prod_name].
    ПОДТВЕРЖДЕНО (Eneba, май): == эталон ЗАГРУЖЕНО, 87/87, Σ|Δ|=0.
    """
    zone = CHANNEL_TO_ZONE.get(channel, channel)
    agg = pipe.aggregate(zone)
    if agg.empty:
        return pd.DataFrame(columns=["ID", "Загружено", "prod_name"])
    out = (agg.groupby("pid", dropna=False)
              .agg(Загружено=("qty", "sum"), prod_name=("prod_name", "first"))
              .reset_index().rename(columns={"pid": "ID"}))
    out["ID"] = pd.to_numeric(out["ID"], errors="coerce").astype("Int64")
    return out


# ---------------------------------------------------------------------------
# Стадия 2 — Каталог: ID -> (base, Регион, Название)
# ---------------------------------------------------------------------------
def catalog_map(catalog_path: str) -> pd.DataFrame:
    """ID, Название, Регион, base. Имя — `Новое назавание` (engine_sales.load_catalog)."""
    cat = sales.load_catalog(catalog_path)            # [ID, Продукт, Новое название, ...]
    cat = cat.copy()
    name_col = "Новое название" if "Новое название" in cat.columns else "Продукт"
    cat["Название"] = cat[name_col].astype(str)
    # Регион — суффикс валюты в названии (если в Каталоге нет явной колонки «Регион»)
    if "Регион" not in cat.columns:
        m = cat["Название"].str.extract(REGION_RE.pattern + "", expand=False)
        cat["Регион"] = m.fillna("").str.upper()
    cat["base"] = cat["Название"].map(_base_name)
    cat["ID"] = pd.to_numeric(cat["ID"], errors="coerce").astype("Int64")
    return cat[["ID", "Название", "Регион", "base"]].dropna(subset=["ID"]).drop_duplicates("ID")


# ---------------------------------------------------------------------------
# Стадия 3 — Остаток = Остаток_нач + Загружено
# ---------------------------------------------------------------------------
def build_stock(zagr: pd.DataFrame, cat: pd.DataFrame,
                ostatok_nach: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """[ID, base, Регион, Название, Остаток_нач, Загружено, Остаток].

    ostatok_nach: [ID, Остаток_нач] — перенос с прошлого месяца. None -> 0.
    """
    df = cat.merge(zagr[["ID", "Загружено"]], on="ID", how="outer")
    df["Загружено"] = pd.to_numeric(df["Загружено"], errors="coerce").fillna(0).astype(int)
    if ostatok_nach is not None and not ostatok_nach.empty:
        on = ostatok_nach.rename(columns={ostatok_nach.columns[-1]: "Остаток_нач"})
        df = df.merge(on[["ID", "Остаток_нач"]], on="ID", how="left")
    else:
        df["Остаток_нач"] = 0
    df["Остаток_нач"] = pd.to_numeric(df["Остаток_нач"], errors="coerce").fillna(0).astype(int)
    df["Остаток"] = df["Остаток_нач"] + df["Загружено"]
    for c in ("base", "Регион", "Название"):
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("")
    return df[["ID", "base", "Регион", "Название", "Остаток_нач", "Загружено", "Остаток"]]


# ---------------------------------------------------------------------------
# Стадия 4 — продажи по каналам (6 парсеров) + спрос по base
# ---------------------------------------------------------------------------
_BUILDERS = {
    "Kinguin": mp_kinguin, "Driffle": mp_driffle, "Eneba": mp_eneba,
    "G2A": mp_g2a, "GGSel": mp_ggsel, "Plati": mp_plati,
}


def sales_by_channel(channel: str, raw_path: str, etalon_path: str,
                     report_date=None) -> pd.DataFrame:
    """Унифицированный лист продаж канала. Маппинг сеется из эталона канала."""
    mod = _BUILDERS[channel]
    if channel == "Kinguin":
        mp = mod.build_kinguin_mapping_from_etalon(etalon_path, raw_path)
        return mod.build_kinguin(raw_path, mp, report_date=report_date)
    if channel == "GGSel":
        seed = mod.build_ggsel_seed(etalon_path)
        return mod.build_ggsel(raw_path, seed, report_date=report_date)
    fn_map = getattr(mod, f"build_{channel.lower()}_mapping_from_etalon")
    fn_build = getattr(mod, f"build_{channel.lower()}")
    mp = fn_map(etalon_path)
    return fn_build(raw_path, mp, report_date=report_date)


def demand_by_base(sheet: pd.DataFrame, cat: pd.DataFrame) -> dict:
    """Спрос по base для регион-агностичных каналов: суммируем Количество по base.

    Лист канала уже несёт ID и Наименование; base берём из Каталога по ID,
    иначе из Наименования.
    """
    s = sheet.merge(cat[["ID", "base"]], on="ID", how="left")
    s["base"] = s["base"].fillna(s["Наименование"].map(_base_name))
    return s.groupby("base")["Количество"].sum().astype(int).to_dict()


# ---------------------------------------------------------------------------
# Стадия 5 — разнесение спроса по региональным ID (для регион-агностичных)
# ---------------------------------------------------------------------------
def allocate_region(stock: pd.DataFrame, demand: dict,
                    region_priority: Optional[list] = None) -> pd.DataFrame:
    """allocator.allocate: спрос(base) + остатки(ID,Регион) -> Разнесено по ID + Флаг."""
    stock_in = stock.rename(columns={"Остаток": "Остаток"})[
        ["ID", "base", "Регион", "Остаток"]]
    return allocator.allocate(stock_in, demand, region_priority=region_priority)


# ---------------------------------------------------------------------------
# Стадия 6 — вычисляемое «Движение ключей»
# ---------------------------------------------------------------------------
def movement(stock: pd.DataFrame, prodano: pd.DataFrame,
             sobytiya: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """нач + загружено − продано + события = конец, по каждому ID.

    prodano:  [ID, Продано]  — свод по всем каналам/регионам.
    sobytiya: [ID, События]  — нетто-ПРИТОК в сток (перемещения-внутрь/возвраты/
              корректировки). ПОДТВЕРЖДЕНО (Eneba/Kinguin/Driffle/G2A, май):
              с этим знаком тождество закрывается 100% на каждой строке каждого
              канала (Σ|Δ|=0). «События» ПРИБАВЛЯЕТСЯ.
    """
    p = prodano.groupby("ID")["Продано"].sum().reset_index()
    mv = stock.merge(p, on="ID", how="left")
    mv["Продано"] = pd.to_numeric(mv["Продано"], errors="coerce").fillna(0).astype(int)
    if sobytiya is not None and not sobytiya.empty:
        ev = sobytiya.groupby("ID")["События"].sum().reset_index()
        mv = mv.merge(ev, on="ID", how="left")
    else:
        mv["События"] = 0
    mv["События"] = pd.to_numeric(mv["События"], errors="coerce").fillna(0).astype(int)
    mv["Остаток_конец"] = mv["Остаток_нач"] + mv["Загружено"] - mv["Продано"] + mv["События"]
    mv["Флаг"] = ""
    mv.loc[mv["Остаток_конец"] < 0, "Флаг"] = "отрицательный остаток"
    return mv[["ID", "Название", "Регион", "base", "Остаток_нач", "Загружено",
               "Продано", "События", "Остаток_конец", "Флаг"]]


# ---------------------------------------------------------------------------
# Гейт валидации — Загружено против эталонного «Движения ключей»
# ---------------------------------------------------------------------------
def validate_zagruzheno(zagr: pd.DataFrame, etalon_movement_path: str,
                        zagr_col_candidates=("Загружено", "Загружено ключей",
                                             "Загружено за период")) -> pd.DataFrame:
    """Сверяет qty-по-pid из KeyFlow с колонкой «Загружено» эталонного «Движения ключей».

    Возврат: таблица [ID, Загружено_KeyFlow, Загружено_эталон, Δ] с расхождениями.
    """
    dk = pd.read_excel(etalon_movement_path, sheet_name="Движение ключей",
                       header=1, engine="calamine")
    dk = dk[pd.to_numeric(dk["ID"], errors="coerce").notna()].copy()
    dk["ID"] = dk["ID"].astype(int)
    col = next((c for c in zagr_col_candidates if c in dk.columns), None)
    if col is None:
        cand = [c for c in dk.columns if "агруж" in str(c)]
        col = cand[0] if cand else None
    if col is None:
        raise KeyError(f"Колонка «Загружено» не найдена. Есть: {list(dk.columns)}")
    et = dk.groupby("ID")[col].sum().reset_index().rename(columns={col: "Загружено_эталон"})
    cmp = zagr.rename(columns={"Загружено": "Загружено_KeyFlow"})[["ID", "Загружено_KeyFlow"]] \
              .merge(et, on="ID", how="outer").fillna(0)
    cmp["Δ"] = cmp["Загружено_KeyFlow"] - cmp["Загружено_эталон"]
    return cmp.sort_values("Δ", key=lambda s: s.abs(), ascending=False)


def build_review(alloc: pd.DataFrame) -> pd.DataFrame:
    """Очередь ревью: структурно-неопределённые позиции (без эталона).

    Флаги allocator ('регион не задан', 'нехватка остатка') + многорегиональные
    базы, где спрос разложен на >=2 региона (кандидаты region-lock/quirk).
    """
    rev = alloc[alloc["Флаг"] != ""].copy()
    placed = alloc[(alloc["ID"].notna()) & (alloc["Разнесено"] > 0)].copy()
    multi = placed.groupby("base")["Регион"].nunique()
    multi_bases = set(multi[multi >= 2].index)
    extra = placed[placed["base"].isin(multi_bases) & (placed["Флаг"] == "")].copy()
    extra["Флаг"] = "спорный сплит (многорегион.)"
    out = pd.concat([rev, extra], ignore_index=True)
    return out.sort_values(["Флаг", "base"]).reset_index(drop=True)


def run_channel(pipe: Pipeline, channel: str, raw_path: str, etalon_path: str,
                cat: pd.DataFrame, ostatok_nach: Optional[pd.DataFrame] = None,
                sobytiya: Optional[pd.DataFrame] = None,
                region_priority: Optional[list] = None,
                report_date=None) -> dict:
    """Полный проход одного канала.

    Возврат: {
      'sales':    унифицированный лист продаж (деньги/итог по листингу),
      'movement': «Движение ключей» по ID (нач+загр−прод+события=конец),
      'alloc':    результат разнесения с флагами,
      'review':   позиции на ревью (флаги + спорный многорегион. сплит),
    }
    """
    sales = sales_by_channel(channel, raw_path, etalon_path, report_date)
    zagr = zagruzheno(pipe, channel)                       # стадия 1 (KeyFlow)
    stock = build_stock(zagr, cat, ostatok_nach)            # стадия 3
    demand = demand_by_base(sales, cat)                     # стадия 4 (из билдера)
    alloc = allocate_region(stock, demand, region_priority) # стадия 5
    prodano = (alloc[alloc["ID"].notna()].groupby("ID")["Разнесено"].sum()
                    .reset_index().rename(columns={"Разнесено": "Продано"}))
    mv = movement(stock, prodano, sobytiya)                 # стадия 6
    return {"sales": sales, "movement": mv, "alloc": alloc,
            "review": build_review(alloc)}


def zagruzheno_zone(pipe: Pipeline, zone: str) -> pd.DataFrame:
    """«Загружено» по pid напрямую из закуп-зоны KeyFlow (для общих реестров)."""
    agg = pipe.aggregate(zone)
    if agg.empty:
        return pd.DataFrame(columns=["ID", "Загружено", "prod_name"])
    out = (agg.groupby("pid", dropna=False)
              .agg(Загружено=("qty", "sum"), prod_name=("prod_name", "first"))
              .reset_index().rename(columns={"pid": "ID"}))
    out["ID"] = pd.to_numeric(out["ID"], errors="coerce").astype("Int64")
    return out


def run_combined(pipe: Pipeline, channels: list, raws: dict, etalon_path: str,
                 cat: pd.DataFrame, ostatok_nach: Optional[pd.DataFrame] = None,
                 sobytiya: Optional[pd.DataFrame] = None, zone: str = "Plati",
                 region_priority: Optional[list] = None, report_date=None) -> dict:
    """Общий реестр (напр. Plati+GGSel): один сток, продажи нескольких каналов.

    Спрос — сумма по всем каналам; сток — из закуп-зоны `zone` (KeyFlow).
    Возврат: {'sales': {канал: лист}, 'movement': общий, 'alloc', 'review'}.
    """
    sales = {}
    dem_total: dict = {}
    for ch in channels:
        s = sales_by_channel(ch, raws[ch], etalon_path, report_date)
        sales[ch] = s
        for b, q in demand_by_base(s, cat).items():
            dem_total[b] = dem_total.get(b, 0) + int(q)
    zagr = zagruzheno_zone(pipe, zone)
    stock = build_stock(zagr, cat, ostatok_nach)
    alloc = allocate_region(stock, dem_total, region_priority)
    prodano = (alloc[alloc["ID"].notna()].groupby("ID")["Разнесено"].sum()
                    .reset_index().rename(columns={"Разнесено": "Продано"}))
    mv = movement(stock, prodano, sobytiya)
    return {"sales": sales, "movement": mv, "alloc": alloc,
            "review": build_review(alloc)}


# ── работа без эталона: стоячий каталог + персистентные маппинги ──
import mp_ggsel as _mp_ggsel  # noqa

_BUILD_FN = {
    "Kinguin": ("mp_kinguin", "build_kinguin"),
    "Driffle": ("mp_driffle", "build_driffle"),
    "Eneba":   ("mp_eneba",   "build_eneba"),
    "G2A":     ("mp_g2a",     "build_g2a"),
    "Plati":   ("mp_plati",   "build_plati"),
    "GGSel":   ("mp_ggsel",   "build_ggsel"),
}


def load_catalog_master(path: str) -> pd.DataFrame:
    """catalog_master.csv -> [ID, base, Регион, Название] для allocator/движения."""
    c = pd.read_csv(path)
    c = c[pd.to_numeric(c["ID"], errors="coerce").notna()].copy()
    c["ID"] = c["ID"].astype("Int64")
    name_col = "Новое назавание" if "Новое назавание" in c.columns else "Продукт"
    c["Название"] = c[name_col].astype(str)
    c["Регион"] = c["Регион"].astype(str) if "Регион" in c.columns else ""
    c["base"] = c["Название"].map(lambda s: REGION_RE.sub("", str(s)).strip())
    return c[["ID", "base", "Регион", "Название"]].drop_duplicates("ID")


def sales_from_mapping(channel: str, raw_path: str, mapping, report_date=None):
    """Унифицированный лист продаж по ПЕРСИСТЕНТНОМУ маппингу (без эталона)."""
    import importlib
    mod_name, fn_name = _BUILD_FN[channel]
    fn = getattr(importlib.import_module(mod_name), fn_name)
    return fn(raw_path, mapping, report_date=report_date)


def run_channel_m(pipe: Pipeline, channel: str, raw_path: str, mapping,
                  cat: pd.DataFrame, ostatok_nach: Optional[pd.DataFrame] = None,
                  sobytiya: Optional[pd.DataFrame] = None,
                  region_priority: Optional[list] = None, report_date=None) -> dict:
    """Как run_channel, но продажи по персистентному маппингу. Новые листинги
    (ID=NA после билдера) выносятся в out['new_listings']."""
    sales = sales_from_mapping(channel, raw_path, mapping, report_date)
    new_listings = sales[sales["ID"].isna()].copy()
    zagr = zagruzheno(pipe, channel)
    stock = build_stock(zagr, cat, ostatok_nach)
    demand = demand_by_base(sales, cat)
    alloc = allocate_region(stock, demand, region_priority)
    prodano = (alloc[alloc["ID"].notna()].groupby("ID")["Разнесено"].sum()
                    .reset_index().rename(columns={"Разнесено": "Продано"}))
    mv = movement(stock, prodano, sobytiya)
    return {"sales": sales, "movement": mv, "alloc": alloc,
            "review": build_review(alloc), "new_listings": new_listings}


def run_combined_m(pipe: Pipeline, channels: list, raws: dict, mappings: dict,
                   cat: pd.DataFrame, ostatok_nach: Optional[pd.DataFrame] = None,
                   sobytiya: Optional[pd.DataFrame] = None, zone: str = "Plati",
                   region_priority: Optional[list] = None, report_date=None,
                   extra_demand: Optional[dict] = None) -> dict:
    """Общий реестр по персистентным маппингам.

    extra_demand: {catalog_base: qty} — доп. спрос (напр. Plati-бандлы,
    разложенные по изданиям), добавляется к спросу билдеров перед allocator.
    """
    sales, dem_total, new_all = {}, {}, []
    for ch in channels:
        s = sales_from_mapping(ch, raws[ch], mappings[ch], report_date)
        sales[ch] = s
        nl = s[s["ID"].isna()].copy()
        if len(nl):
            nl.insert(0, "Канал", ch); new_all.append(nl)
        for b, q in demand_by_base(s, cat).items():
            dem_total[b] = dem_total.get(b, 0) + int(q)
    if extra_demand:
        for b, q in extra_demand.items():
            dem_total[b] = dem_total.get(b, 0) + int(q)
    zagr = zagruzheno_zone(pipe, zone)
    stock = build_stock(zagr, cat, ostatok_nach)
    alloc = allocate_region(stock, dem_total, region_priority)
    prodano = (alloc[alloc["ID"].notna()].groupby("ID")["Разнесено"].sum()
                    .reset_index().rename(columns={"Разнесено": "Продано"}))
    mv = movement(stock, prodano, sobytiya)
    new_listings = pd.concat(new_all, ignore_index=True) if new_all else pd.DataFrame()
    return {"sales": sales, "movement": mv, "alloc": alloc,
            "review": build_review(alloc), "new_listings": new_listings}


def split_sales_by_alloc(sales: pd.DataFrame, alloc: pd.DataFrame,
                         cat: pd.DataFrame) -> pd.DataFrame:
    """[DEPRECATED — НЕ ИСПОЛЬЗОВАТЬ ДЛЯ QB] Разбивка листа продаж по регионам.

    ВНИМАНИЕ: теряет выручку (~33% на Kinguin), т.к. allocator разносит спрос
    только до наличия стока — неразнесённый остаток выпадает. Для QB-листа
    продаж выручка ДОЛЖНА сохраняться точно → используется ЛУМП-лист (sales).
    Региональное измерение для COGS берётся из ledger движения, а не отсюда.
    Оставлено только для экспериментов; в пайплайн не подключено.
    """
    cat_base = cat.set_index("ID")["base"].to_dict()
    name = cat.set_index("ID")["Название"].to_dict()
    reg = cat.set_index("ID")["Регион"].to_dict()
    # удельная цена и валюта по base из sales
    s = sales.copy()
    s["base"] = s["ID"].map(cat_base).fillna(s["Наименование"])
    agg = s.groupby("base").agg(qty=("Количество", "sum"), summ=("Сумма", "sum"),
                                cur=("Валюта", "first"),
                                date=("Дата", "first")).reset_index()
    agg["unit"] = (agg["summ"] / agg["qty"]).where(agg["qty"] != 0, 0.0)
    unit = dict(zip(agg["base"], agg["unit"]))
    cur = dict(zip(agg["base"], agg["cur"]))
    date = dict(zip(agg["base"], agg["date"]))
    rows = []
    placed = alloc[alloc["ID"].notna() & (alloc["Разнесено"] > 0)]
    for r in placed.itertuples():
        b = r.base
        u = unit.get(b, 0.0)
        rows.append({"Дата": date.get(b, pd.NaT), "ID": int(r.ID),
                     "Наименование": name.get(int(r.ID), b),
                     "Партнер": "", "Количество": int(r.Разнесено),
                     "Цена": round(u, 6), "Валюта": cur.get(b, ""),
                     "Сумма": round(r.Разнесено * u, 2)})
    out = pd.DataFrame(rows, columns=["Дата", "ID", "Наименование", "Партнер",
                                      "Количество", "Цена", "Валюта", "Сумма"])
    # добавить несопоставленные (ID=NA) из исходных продаж как есть
    unm = sales[sales["ID"].isna()]
    return pd.concat([out, unm], ignore_index=True)


# ===========================================================================
# ТОНКИЙ ОРКЕСТРАТОР «ПРОГНАТЬ ВСЁ» (закуп + продажи за один проход биллинга)
# ---------------------------------------------------------------------------
# Один загруженный Pipeline (R1/R2/genba, ~40с) кормит ОБЕ стороны:
#   ЗАКУП   — KeyFlow по 9 площадкам (aggregate -> to_dataframe = QB-свод);
#   ПРОДАЖИ — SalesFlow по 6 каналам (run_channel_m / run_combined_m).
# Здесь только КОМПОЗИЦИЯ существующих функций — никакой новой бизнес-логики.
# Гейт сверки: Загружено(движения) == Σ qty(закуп-агрегата) по зоне канала.
# ===========================================================================

# Все площадки закупа (KeyFlow). Порядок — как в конфиге.
PURCHASE_ZONES = list(PLOSHADKA_MAP.keys())   # Plati,Kinguin,Eneba,G2A,Driffle,Tao,ChinaPlay,B2B,GamersBase

# Индивидуальные каналы продаж (регион разносится allocator по своей закуп-зоне).
SALES_INDIV = ["Eneba", "Kinguin", "Driffle", "G2A"]

# Юнит движения -> закуп-зона, которая формирует его «Загружено» (для гейта).
# GGSel не имеет своей закуп-зоны: Plati+GGSel делят зону «Plati» (см. память).
MOVEMENT_TO_ZONE = {
    "Eneba": "Eneba", "Kinguin": "Kinguin", "Driffle": "Driffle",
    "G2A": "G2A", "Plati+GGSel": "Plati",
}


# --- ЗАКУП: тонкая обёртка над Pipeline.aggregate + to_dataframe -----------
def run_purchases(pipe: Pipeline, zones: Optional[list] = None,
                  active_only: bool = True) -> dict:
    """Отчёт закупа по каждой площадке из одного загруженного Pipeline.

    Возврат: {zone: {'agg': сырой агрегат, 'flat': QB-свод (to_dataframe),
                     'qty_agg': Σqty сырого агрегата (базис гейта),
                     'qty': Σqty в своде (active), 'cost': Σсебестоимость}}.
    Пустые площадки пропускаются. Никакой логики — только вызовы движка.
    """
    zones = zones or PURCHASE_ZONES
    out = {}
    for z in zones:
        agg = pipe.aggregate(z)
        if agg is None or agg.empty:
            continue
        flat = pipe.to_dataframe(agg, z, active_only=active_only)
        act = agg[agg["qty"] > 0]
        # строки, которые to_dataframe роняет: qty>0, но поставщик не классифицирован.
        # Их НЕ теряем — отдаём отдельно с фолбэком supp_raw для свода/аннотаций.
        unc_src = act[act["supplier_group"].isna()]
        unclassified = pd.DataFrame({
            "Площадка": z,
            "Поставщик": unc_src["supp_raw"].fillna("(не указан)"),
            "ID продукта": unc_src["pid"].astype("Int64"),
            "Название": unc_src["prod_name"].fillna(""),
            "Количество": unc_src["qty"].astype("Int64"),
            "Цена закупа": unc_src["unit_price"].round(4),
            "Валюта": unc_src["currency"],
            "Себестоимость": unc_src["cost"].round(2),
        }).reset_index(drop=True)
        out[z] = {
            "agg": agg,
            "flat": flat,
            "unclassified": unclassified,               # qty>0 с NaN-поставщиком
            "qty_agg": int(agg["qty"].sum()),          # тот же базис, что zagruzheno
            "qty": int(act["qty"].sum()),              # то, что видно в своде
            "cost": float(act["cost"].sum()),
            "suppliers": int(act["supplier_group"].nunique()),
        }
    return out


def purchases_flat(purchases: dict) -> pd.DataFrame:
    """Единый плоский лист «Закуп (все площадки)» — конкатенация сводов to_dataframe.

    Колонки: [Площадка, Поставщик, ID продукта, Название, Количество,
              Цена закупа, Валюта, Себестоимость].
    """
    frames = [v["flat"] for v in purchases.values()
              if isinstance(v.get("flat"), pd.DataFrame) and len(v["flat"])]
    if not frames:
        return pd.DataFrame(columns=["Площадка", "Поставщик", "ID продукта",
                                     "Название", "Количество", "Цена закупа",
                                     "Валюта", "Себестоимость"])
    return pd.concat(frames, ignore_index=True)


# --- ПРОДАЖИ: единый прогон 6 каналов (то, что раньше жило в app.py) --------
def run_sales_all(pipe: Pipeline, cat: pd.DataFrame, raws: dict, mappings: dict,
                  carry: Optional[dict] = None, events=None,
                  report_date=None, region_priority: Optional[list] = None,
                  extra_demand: Optional[dict] = None,
                  errors: Optional[list] = None) -> dict:
    """Прогон всех каналов продаж по персистентным маппингам (без эталона).

    raws:     {канал: путь_к_выгрузке}   (Eneba/Kinguin/Driffle/G2A/Plati/GGSel)
    mappings: {канал: маппинг}           (загруженный extract_mappings.load_*)
    carry:    {канал|'Plati+GGSel'|'*': DataFrame[ID, Остаток_нач]}
    extra_demand: {base: qty} — доп. спрос (Plati-бандлы), в общий реестр.

    Возврат: {канал: run_channel_m(...)} + при наличии Plati&GGSel —
    {'Plati+GGSel': {sales_multi, movement, review, new_listings}}.
    Структура совпадает с той, что ждёт app.py.
    """
    carry = carry or {}

    def _carry(key):
        return carry.get(key, carry.get("*"))

    # events: либо общий DataFrame (на все каналы), либо dict {канал|'*': df}.
    def _events(key):
        if isinstance(events, dict):
            return events.get(key, events.get("*"))
        return events

    results = {}
    for ch in SALES_INDIV:
        if raws.get(ch) is None or mappings.get(ch) is None:
            continue
        try:
            results[ch] = run_channel_m(
                pipe, ch, raws[ch], mappings[ch], cat,
                ostatok_nach=_carry(ch), sobytiya=_events(ch),
                region_priority=region_priority, report_date=report_date)
        except Exception as e:  # изоляция: сбой канала не рушит остальные
            if errors is not None:
                errors.append((ch, f"{type(e).__name__}: {e}"))
            else:
                raise

    if raws.get("Plati") and raws.get("GGSel"):
        try:
            out = run_combined_m(
                pipe, ["Plati", "GGSel"],
                {"Plati": raws["Plati"], "GGSel": raws["GGSel"]},
                {"Plati": mappings.get("Plati"), "GGSel": mappings.get("GGSel")},
                cat, ostatok_nach=_carry("Plati+GGSel"), sobytiya=_events("Plati+GGSel"),
                zone="Plati", region_priority=region_priority,
                report_date=report_date, extra_demand=extra_demand)
            results["Plati+GGSel"] = {
                "sales_multi": out["sales"], "movement": out["movement"],
                "review": out["review"], "new_listings": out["new_listings"]}
        except Exception as e:
            if errors is not None:
                errors.append(("Plati+GGSel", f"{type(e).__name__}: {e}"))
            else:
                raise
    return results


# --- ГЕЙТ СВЕРКИ: Загружено(движения) ↔ Σqty(закуп-зоны) -------------------
def reconcile(purchases: dict, results: dict) -> pd.DataFrame:
    """Сверяет «Загружено» каждого движения с Σqty закуп-агрегата его зоны.

    По построению Δ должна быть 0 (один и тот же aggregate). Ненулевая Δ —
    сигнал рассинхрона проводки. Площадки закупа без движения помечаются
    отдельно (продажи для них в этом проходе не строятся).
    Возврат: [Юнит, Зона, Загружено_движение, Закуп_qty, Δ, Статус].
    """
    rows = []
    covered_zones = set()
    for unit, zone in MOVEMENT_TO_ZONE.items():
        if unit not in results:
            continue
        mv = results[unit]["movement"]
        zagr_mv = int(pd.to_numeric(mv["Загружено"], errors="coerce").fillna(0).sum())
        p = purchases.get(zone)
        zak = int(p["qty_agg"]) if p else 0
        covered_zones.add(zone)
        d = zagr_mv - zak
        rows.append({"Юнит": unit, "Зона": zone,
                     "Загружено_движение": zagr_mv, "Закуп_qty": zak,
                     "Δ": d, "Статус": "OK" if d == 0 else "РАСХОЖДЕНИЕ"})
    for zone, p in purchases.items():
        if zone in covered_zones:
            continue
        rows.append({"Юнит": "—", "Зона": zone,
                     "Загружено_движение": 0, "Закуп_qty": int(p["qty_agg"]),
                     "Δ": 0, "Статус": "закуп без движения (продажи не строятся)"})
    return pd.DataFrame(rows, columns=["Юнит", "Зона", "Загружено_движение",
                                       "Закуп_qty", "Δ", "Статус"])


# --- ВЕРХНЯЯ КОМПОЗИЦИЯ: прогнать всё за один проход ------------------------
def run_everything(pipe: Pipeline, cat: pd.DataFrame, raws: dict, mappings: dict,
                   carry: Optional[dict] = None, events: Optional[pd.DataFrame] = None,
                   report_date=None, purchase_zones: Optional[list] = None,
                   region_priority: Optional[list] = None,
                   extra_demand: Optional[dict] = None) -> dict:
    """Единый проход: закуп (9 площадок) + продажи/движение (6 каналов) + гейт.

    Pipeline загружается ВНЕ функции (один раз, ~40с) и передаётся сюда —
    так закуп и продажи делят одну загрузку биллинга.

    Возврат: {
      'purchases':      {zone: {...}}            — закуп по площадкам,
      'purchases_flat': DataFrame                — единый плоский лист закупа,
      'sales':          {канал: run_channel_m}   — продажи+движение,
      'reconcile':      DataFrame                — гейт сверки Загружено↔закуп,
    }
    """
    purchases = run_purchases(pipe, purchase_zones)
    errors: list = []
    results = run_sales_all(pipe, cat, raws, mappings, carry=carry, events=events,
                            report_date=report_date, region_priority=region_priority,
                            extra_demand=extra_demand, errors=errors)
    rec = reconcile(purchases, results)
    return {"purchases": purchases,
            "purchases_flat": purchases_flat(purchases),
            "sales": results,
            "reconcile": rec,
            "review": review_consolidated(results, catalog=cat),
            "errors": errors}


# --- КОНСОЛИДИРОВАННЫЙ ЛИСТ РУЧНОЙ СВЕРКИ ----------------------------------
# Объединяет per-channel `review` (хвост аллокатора) и `new_listings`
# (листинги без catID) в один лист с причиной, категорией действия и
# нераспределённым спросом. Новой логики нет — только сбор существующих
# артефактов + типизация для глаз человека.
REVIEW_CATEGORY = {
    "листинг не в маппинге":         "Сопоставить листинг",
    "регион не задан":               "Задать регион (склад)",
    "спорный сплит (многорегион.)":  "Задать регион (склад)",
    "нехватка остатка":              "Вероятно снимется carry-over",
}
_CAT_ORDER = {"Сопоставить листинг": 0, "Задать регион (склад)": 1,
              "Вероятно снимется carry-over": 2, "Прочее": 9}
_REVIEW_COLS = ["Канал", "Источник", "Категория", "Причина", "Позиция", "ID",
                "Регион", "Остаток", "Спрос", "Разнесено", "Нераспределено"]


def review_consolidated(sales_results: dict, catalog: "pd.DataFrame | None" = None) -> pd.DataFrame:
    """Единый лист ручной сверки по всем каналам продаж.

    Вход: {канал: результат run_channel_m/run_combined_m} (у каждого есть
    `review` и `new_listings`). Если передан catalog (ID, base, Регион, …),
    к строкам «регион не задан» добавляется колонка «Регионы-кандидаты» —
    список catID(Регион) с тем же base, чтобы аналитик знал варианты выбора.
    """
    frames = []
    for ch, r in (sales_results or {}).items():
        rv = r.get("review")
        if isinstance(rv, pd.DataFrame) and len(rv):
            t = rv.copy()
            t["Канал"] = ch
            t["Источник"] = "allocator"
            t["Причина"] = t["Флаг"] if "Флаг" in t.columns else ""
            t = t.rename(columns={"base": "Позиция"})
            for c in ("ID", "Регион", "Остаток", "Спрос", "Разнесено"):
                if c not in t.columns:
                    t[c] = 0 if c in ("Остаток", "Спрос", "Разнесено") else "—"
            frames.append(t[["Канал", "Источник", "Причина", "Позиция", "ID",
                             "Регион", "Остаток", "Спрос", "Разнесено"]])
        nl = r.get("new_listings")
        if isinstance(nl, pd.DataFrame) and len(nl):
            frames.append(pd.DataFrame({
                "Канал": ch, "Источник": "listing",
                "Причина": "листинг не в маппинге",
                "Позиция": nl.get("Наименование"), "ID": nl.get("ID"),
                "Регион": "—", "Остаток": 0,
                "Спрос": pd.to_numeric(nl.get("Количество"), errors="coerce").fillna(0),
                "Разнесено": 0}))
    if not frames:
        return pd.DataFrame(columns=_REVIEW_COLS)
    out = pd.concat(frames, ignore_index=True)
    out["Нераспределено"] = (pd.to_numeric(out["Спрос"], errors="coerce").fillna(0)
                             - pd.to_numeric(out["Разнесено"], errors="coerce").fillna(0)).astype(int)
    out["Категория"] = out["Причина"].map(REVIEW_CATEGORY).fillna("Прочее")
    out["_o"] = out["Категория"].map(_CAT_ORDER).fillna(9)
    out = (out.sort_values(["_o", "Нераспределено"], ascending=[True, False])
              .drop(columns="_o").reset_index(drop=True))
    cols = list(_REVIEW_COLS)
    if catalog is not None and "base" in catalog.columns:
        cand = (catalog.dropna(subset=["base"])
                .assign(_c=lambda d: d["ID"].astype("Int64").astype(str)
                        + "(" + d["Регион"].fillna("?").astype(str) + ")")
                .groupby("base")["_c"].apply(lambda s: ", ".join(s.astype(str).head(8))).to_dict())
        need = out["Причина"].eq("регион не задан")
        out["Регионы-кандидаты"] = ""
        out.loc[need, "Регионы-кандидаты"] = out.loc[need, "Позиция"].map(cand).fillna("")
        cols = cols + ["Регионы-кандидаты"]
    return out[cols]
