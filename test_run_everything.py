"""Smoke-тесты тонкого оркестратора «прогнать всё» (закуп + продажи).

Запуск (без биллинга — проверяется композиция и гейт на синтетике):
    PYTHONPATH=. python3 test_run_everything.py

Логика движков (KeyFlow.aggregate, run_channel_m) уже проверена на данных
мая 2026 в отдельных сессиях; здесь проверяется ТОЛЬКО тонкий слой:
run_purchases / purchases_flat / reconcile / run_everything.
"""
from __future__ import annotations

import pandas as pd

import orchestrator as o
from engine_keyflow import Pipeline


def test_run_purchases_empty():
    """Пустой Pipeline (нет биллинга) -> пустой закуп, без падений."""
    assert o.run_purchases(Pipeline()) == {}


def test_purchases_flat_concat():
    """Плоский лист = конкатенация сводов to_dataframe по всем площадкам."""
    fake = {
        "Eneba": {"flat": pd.DataFrame({
            "Площадка": ["Eneba", "Eneba"], "Поставщик": ["A", "A"],
            "ID продукта": [1, 2], "Название": ["g1", "g2"],
            "Количество": [10, 5], "Цена закупа": [1.0, 2.0],
            "Валюта": ["EUR", "EUR"], "Себестоимость": [10.0, 10.0]})},
        "B2B": {"flat": pd.DataFrame({
            "Площадка": ["B2B"], "Поставщик": ["C"], "ID продукта": [4],
            "Название": ["g4"], "Количество": [3], "Цена закупа": [5.0],
            "Валюта": ["USD"], "Себестоимость": [15.0]})},
    }
    flat = o.purchases_flat(fake)
    assert len(flat) == 3
    assert set(flat["Площадка"]) == {"Eneba", "B2B"}


def test_purchases_flat_empty():
    assert list(o.purchases_flat({}).columns) == [
        "Площадка", "Поставщик", "ID продукта", "Название", "Количество",
        "Цена закупа", "Валюта", "Себестоимость"]


def test_reconcile_gate():
    """Гейт: Загружено(движение) vs Σqty(закуп-зоны). OK / РАСХОЖДЕНИЕ / без движения."""
    purchases = {
        "Eneba": {"qty_agg": 15, "flat": pd.DataFrame()},
        "Plati": {"qty_agg": 7, "flat": pd.DataFrame()},
        "B2B": {"qty_agg": 3, "flat": pd.DataFrame()},
    }
    results = {
        "Eneba": {"movement": pd.DataFrame({"ID": [1, 2], "Загружено": [10, 5]})},
        "Plati+GGSel": {"movement": pd.DataFrame({"ID": [3], "Загружено": [6]})},
    }
    rec = o.reconcile(purchases, results).set_index("Зона")
    assert rec.loc["Eneba", "Статус"] == "OK" and rec.loc["Eneba", "Δ"] == 0
    assert rec.loc["Plati", "Статус"] == "РАСХОЖДЕНИЕ" and rec.loc["Plati", "Δ"] == -1
    assert "закуп без движения" in rec.loc["B2B", "Статус"]


def test_review_consolidated():
    """Единый лист: категория действия + сортировка по нераспределённому."""
    sales = {
        "Eneba": {
            "review": pd.DataFrame({
                "base": ["A", "B"], "ID": [1, pd.NA], "Регион": ["EU", "—"],
                "Остаток": [0, 0], "Спрос": [10, 200], "Разнесено": [10, 0],
                "Флаг": ["регион не задан", "нехватка остатка"]}),
            "new_listings": pd.DataFrame({
                "Наименование": ["NewGame"], "ID": [pd.NA], "Количество": [5]}),
        },
    }
    rv = o.review_consolidated(sales)
    assert list(rv.columns) == o._REVIEW_COLS
    # "Сопоставить листинг" идёт первым (действеннее carry-шума)
    assert rv.iloc[0]["Категория"] == "Сопоставить листинг"
    # carry-категория внизу
    assert rv.iloc[-1]["Категория"] == "Вероятно снимется carry-over"
    assert int(rv["Нераспределено"].sum()) == 5 + 0 + 200  # 205


def test_review_consolidated_empty():
    assert list(o.review_consolidated({}).columns) == o._REVIEW_COLS
    assert len(o.review_consolidated({})) == 0


def test_run_everything_wiring():
    """run_everything возвращает 5 ключей и не падает на пустых входах."""
    res = o.run_everything(
        Pipeline(),
        pd.DataFrame(columns=["ID", "base", "Регион", "Название"]),
        raws={}, mappings={})
    assert set(res.keys()) == {"purchases", "purchases_flat", "sales",
                               "reconcile", "review"}
    assert res["purchases"] == {}
    assert res["sales"] == {}
    assert len(res["review"]) == 0


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print(f"  OK  {t.__name__}")
    print(f"\n{len(tests)}/{len(tests)} тестов пройдено.")
