"""
SalesFlow — реестровый сплиттер Plati/ggsel по изданиям.

Издания/регионы Plati берутся НЕ из строк выгрузки, а из реестра «Движение ключей»
(по каталоговым ID) + «Каталог» (имя/регион через 'Новое назавание' по столбцу 'по').

Лист «Движение ключей» имеет двухуровневую шапку. Колонки продаж по каналам
(по позиции; имена 'Продано *' повторяются в двух группах PLATI / PLATI NUKEM):

  col7  ggsel            -> лист 'ggsel'
  col8  Продано WMT      -> PLATIWMT
  col9  Продано RUB      -> PLATIRUB
  col10 Продано WMZ      -> PLATIWMZ NUKEN   (USD)
  col11 Продано WMT      -> PLATIWMT NUKEN   (USD)
  col12 Продано RUB      -> PLATIRUB NUKEN   (RUB)
  col13 Продано WMZ_3    -> PLATI_3_WMZ
  col14 Продано WMT_3    -> PLATI_3_WMT

Канал -> валюта: *WMZ/*WMT -> USD, *RUB -> RUB, ggsel -> USD.
"""
from __future__ import annotations
import pandas as pd
from typing import Optional

# позиция колонки в листе -> (имя канала, лист, валюта)
CHANNELS = {
    7:  ("ggsel",        "ggsel",          "USD"),
    8:  ("PLATI WMT",    "PLATIWMT",       "USD"),
    9:  ("PLATI RUB",    "PLATIRUB",       "RUB"),
    10: ("NUKEM WMZ",    "PLATIWMZ NUKEN", "USD"),
    11: ("NUKEM WMT",    "PLATIWMT NUKEN", "USD"),
    12: ("NUKEM RUB",    "PLATIRUB NUKEN", "RUB"),
    13: ("WMZ_3",        "PLATI_3_WMZ",    "USD"),
    14: ("WMT_3",        "PLATI_3_WMT",    "USD"),
}
PARTNER_BY_CUR = {"RUB": "Физическое лицо PLATI rub", "USD": "Физическое лицо PLATI USD"}


def load_catalog(etalon_path) -> pd.DataFrame:
    cat = pd.read_excel(etalon_path, sheet_name="Каталог", header=0, engine="calamine")
    cat = cat[pd.to_numeric(cat["ID"], errors="coerce").notna()].copy()
    cat["ID"] = cat["ID"].astype(int)
    # 'Новое назавание' — каноничное имя с регионом
    name_col = "Новое назавание" if "Новое назавание" in cat.columns else "Продукт"
    return cat[["ID", name_col]].rename(columns={name_col: "Название"}).drop_duplicates("ID")


def load_movement(etalon_path) -> pd.DataFrame:
    """Читает 'Движение ключей' с двухуровневой шапкой, возвращает ID + колонки каналов по позиции."""
    dk = pd.read_excel(etalon_path, sheet_name="Движение ключей", header=1, engine="calamine")
    dk = dk[pd.to_numeric(dk["ID"], errors="coerce").notna()].copy()
    dk["ID"] = dk["ID"].astype(int)
    return dk


def build_plati_editions(etalon_path,
                         channels: Optional[list] = None,
                         report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """Возвращает продажи по изданиям из реестра. channels — позиции колонок (по умолчанию все)."""
    cat = load_catalog(etalon_path)
    dk = load_movement(etalon_path)
    cat_map = cat.set_index("ID")["Название"]

    use = channels or list(CHANNELS)
    rows = []
    for pos in use:
        chname, sheet, cur = CHANNELS[pos]
        qty = pd.to_numeric(dk.iloc[:, pos], errors="coerce").fillna(0)
        sub = pd.DataFrame({"ID": dk["ID"], "Кол-во": qty})
        sub = sub[sub["Кол-во"] != 0]
        if sub.empty:
            continue
        sub["Канал"] = chname; sub["Лист"] = sheet; sub["Валюта"] = cur
        rows.append(sub)
    if not rows:
        return pd.DataFrame(columns=["ID","Название","Канал","Лист","Валюта","Количество"])
    out = pd.concat(rows, ignore_index=True)
    out["Название"] = out["ID"].map(cat_map)
    out["Партнер"] = out["Валюта"].map(PARTNER_BY_CUR)
    out["Дата"] = pd.to_datetime(report_date) if report_date is not None else pd.NaT
    out = out.rename(columns={"Кол-во": "Количество"})
    return out[["Дата","Лист","ID","Название","Канал","Партнер","Количество","Валюта"]] \
             .sort_values(["Лист","ID"]).reset_index(drop=True)
