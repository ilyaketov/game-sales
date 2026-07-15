"""
SalesFlow — движок разнесения ключей (allocator).

Задача: продажу игры на витрине (спрос по «игра + издание») разложить по
конкретным региональным каталоговым ID, опираясь на остаток ключей.

Правило (выведено из эталона, ~90% игр воспроизводятся точно):
  спрос распределяется по регионам в порядке дешевизны (region_priority),
  каждому региону отдаётся весь доступный остаток (Остаток_нач + Загружено),
  пока спрос не исчерпан.

Флаги на ревью:
  - 'нехватка остатка'   : спроса больше, чем суммарный остаток по всем регионам;
  - 'регион не задан'    : у позиции пустой/неизвестный регион;
  - 'остаток не выбран'  : регион дешевле, но остался непроданный остаток
                           (спрос кончился раньше) — нормально, не ошибка.

Вход:
  stock_df  — DataFrame: [ID, base, Регион, Остаток] (Остаток = нач + загружено).
  demand    — dict {base: спрос_штук} ИЛИ DataFrame [base, Спрос].
Выход:
  DataFrame: [base, ID, Регион, Остаток, Разнесено, Флаг].
"""
from __future__ import annotations
import pandas as pd
from typing import Optional, Union

# Порядок дешевизны регионов (дешёвый -> дорогой). Настраиваемо; можно
# пересчитать из цены закупа в USD через region_priority_from_costs().
DEFAULT_REGION_PRIORITY = [
    "RUB","RU","UAH","KZT","TRY","TL","ARS","BRL","MXN","INR","CNY",
    "PLN","JPY","KRW","NOK","SEK","DKK","EUR","USD","GBP","AUD","CAD","CHF",
]


def region_priority_from_costs(cost_usd: dict) -> list:
    """Строит порядок регионов из цены закупа в USD (дешёвый -> дорогой)."""
    return [r for r, _ in sorted(cost_usd.items(), key=lambda kv: kv[1])]


def _rank(region: str, priority: list) -> int:
    return priority.index(region) if region in priority else len(priority) + 1


def allocate(stock_df: pd.DataFrame,
             demand: Union[dict, pd.DataFrame],
             region_priority: Optional[list] = None,
             stock_col: str = "Остаток",
             base_col: str = "base",
             region_col: str = "Регион",
             id_col: str = "ID") -> pd.DataFrame:
    prio = region_priority or DEFAULT_REGION_PRIORITY
    if isinstance(demand, pd.DataFrame):
        demand = dict(zip(demand[base_col], demand["Спрос"]))

    out = []
    for base, grp in stock_df.groupby(base_col, sort=False):
        D = int(demand.get(base, 0))
        g = grp.copy()
        g["_rank"] = g[region_col].astype(str).map(lambda r: _rank(r, prio))
        g = g.sort_values(["_rank", id_col])
        total_stock = int(pd.to_numeric(g[stock_col], errors="coerce").fillna(0).sum())
        rem = D
        for _, row in g.iterrows():
            st = int(pd.to_numeric(row[stock_col], errors="coerce") or 0)
            take = min(rem, st) if rem > 0 else 0
            rem -= take
            flag = ""
            if str(row[region_col]).strip() in ("", "nan", "None"):
                flag = "регион не задан"
            elif take < st and rem <= 0 and take >= 0:
                flag = ""  # нормальный недобор: спрос кончился
            out.append({
                "base": base, "ID": row[id_col], "Регион": row[region_col],
                "Остаток": st, "Спрос": D, "Разнесено": take, "Флаг": flag,
            })
        if rem > 0:  # спрос не уместился в остатки
            out.append({"base": base, "ID": pd.NA, "Регион": "—",
                        "Остаток": 0, "Спрос": D, "Разнесено": rem,
                        "Флаг": "нехватка остатка"})
    return pd.DataFrame(out)
