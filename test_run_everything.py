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
                               "reconcile", "review", "errors"}
    assert res["purchases"] == {}
    assert res["sales"] == {}
    assert res["errors"] == []
    assert len(res["review"]) == 0


def _mk_xlsx(tmp, name, columns, sheet="Sheet1"):
    import pandas as pd
    p = f"{tmp}/{name}"
    pd.DataFrame(columns=columns).to_excel(p, index=False, sheet_name=sheet)
    return p


def test_detect_cabinets_and_billing(tmp_path=None):
    """Детектор распознаёт кабинеты и биллинг по сигнатуре колонок."""
    import tempfile, detect, config
    tmp = tempfile.mkdtemp()
    cases = {
        "Kinguin": ["PRODUCTID", "PRICE", "QTY"],
        "Driffle": ["Product Title", "Selling Price (EUR)"],
        "G2A": ["Type", "Amount EUR (Approx,)", "Qty"],
        "Eneba": ["Тип", "Оплаченная сумма", "name"],
        "Plati": ["зачислено", "название товара"],
        "r1": list(config.COLS_R1.values()),
        "r2": list(config.COLS_R2.values()),
        "genba": list(config.COLS_GENBA.values()),
    }
    for kind, cols in cases.items():
        p = _mk_xlsx(tmp, f"{kind}.xlsx", cols)
        assert detect.detect_kind(p) == kind, f"{kind}: got {detect.detect_kind(p)}"


def test_detect_carry_events_and_unknown():
    import tempfile, detect, pandas as pd
    tmp = tempfile.mkdtemp()
    pd.DataFrame({"Канал": [], "ID": [], "Остаток_конец": []}).to_csv(f"{tmp}/c.csv", index=False)
    pd.DataFrame({"ID": [], "События": []}).to_csv(f"{tmp}/e.csv", index=False)
    assert detect.detect_kind(f"{tmp}/c.csv") == "carry"
    assert detect.detect_kind(f"{tmp}/e.csv") == "events"
    # многоколоночная книга не должна цепляться как carry/events
    p = _mk_xlsx(tmp, "big.xlsx", ["ID", "События", "A", "B", "C", "D", "E", "F"])
    assert detect.detect_kind(p) is None


def test_report_xlsx_build():
    """Аннотированная книга собирается и открывается; подсветка неклассиф. поставщика."""
    import io, report_xlsx, openpyxl
    res = {
        "mode": "Прогнать всё",
        "reconcile": pd.DataFrame({"Юнит": ["Eneba", "—"], "Зона": ["Eneba", "B2B"],
                                   "Загружено_движение": [10, 0], "Закуп_qty": [10, 5],
                                   "Δ": [0, 0], "Статус": ["OK", "закуп без движения"]}),
        "purchases": {"Eneba": {
            "flat": pd.DataFrame({"Площадка": ["Eneba"], "Поставщик": ["ROKKY"],
                "ID продукта": [1], "Название": ["g1"], "Количество": [10],
                "Цена закупа": [1.0], "Валюта": ["EUR"], "Себестоимость": [10.0]}),
            "unclassified": pd.DataFrame({"Площадка": ["Eneba"], "Поставщик": ["Moogold"],
                "ID продукта": [2], "Название": ["g2"], "Количество": [5],
                "Цена закупа": [2.0], "Валюта": ["USD"], "Себестоимость": [10.0]})}},
        "purchases_flat": pd.DataFrame(),
        "sales": {"Eneba": {
            "sales": pd.DataFrame({"ID": [1], "Наименование": ["g1"], "Количество": [3],
                                   "Цена": [2.0], "Валюта": ["EUR"], "Сумма": [6.0]}),
            "movement": pd.DataFrame({"ID": [1], "Название": ["g1"], "Регион": ["EU"],
                "Остаток_нач": [0], "Загружено": [10], "Продано": [12], "События": [0],
                "Остаток_конец": [-2]})}},
        "review": pd.DataFrame({"Канал": ["Eneba"], "Источник": ["allocator"],
            "Категория": ["Задать регион (склад)"], "Причина": ["регион не задан"],
            "Позиция": ["g1"], "ID": [1], "Регион": ["—"], "Остаток": [0],
            "Спрос": [5], "Разнесено": [0], "Нераспределено": [5]}),
    }
    data = report_xlsx.build(res)
    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert "Легенда и проблемы" in wb.sheetnames
    assert "Закуп (свод)" in wb.sheetnames
    # в своде должна быть подсвеченная (оранжевая) ячейка неклассиф. поставщика
    sv = wb["Закуп (свод)"]
    oranges = [c for row in sv.iter_rows(min_row=2) for c in row
               if c.fill and c.fill.fgColor and str(c.fill.fgColor.rgb)[-6:] == "FFC9A3"]
    assert len(oranges) >= 1 and any(c.comment for c in oranges)


def test_review_candidates():
    """С каталогом строки «регион не задан» получают колонку кандидатов."""
    sales = {"Eneba": {"review": pd.DataFrame({
        "base": ["Game A"], "ID": [pd.NA], "Регион": ["—"], "Остаток": [0],
        "Спрос": [10], "Разнесено": [0], "Флаг": ["регион не задан"]})}}
    cat = pd.DataFrame({"ID": [101, 102], "base": ["Game A", "Game A"],
                        "Регион": ["RUB", "EU"], "Название": ["Game A RUB", "Game A EU"]})
    rv = o.review_consolidated(sales, catalog=cat)
    assert "Регионы-кандидаты" in rv.columns
    cell = rv.iloc[0]["Регионы-кандидаты"]
    assert "101(RUB)" in cell and "102(EU)" in cell


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items())
             if k.startswith("test_") and callable(v)]
    for t in tests:
        t()
        print(f"  OK  {t.__name__}")
    print(f"\n{len(tests)}/{len(tests)} тестов пройдено.")
