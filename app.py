"""
SalesFlow — Streamlit-приложение для сборки отчётов продаж по B2B,
Chinaplay и GamersBase из биллингового Universal Report.

Опционально: загрузка файлов с каталогами для маппинга
билинговых ID на каталоговые.
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from engine import build_all, load_catalog, load_partner_catalog, summary


# ---------- helpers ----------------------------------------------------------

def _to_xlsx_bytes(reports: dict) -> bytes:
    """Записывает три отчёта в один xlsx с именами листов как в эталонах."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        reports["B2B"].to_excel(writer, sheet_name="Продажи Б2Б", index=False)
        reports["Chinaplay"].to_excel(writer, sheet_name="продажи_Чайна", index=False)
        # ВАЖНО: имя листа GB — с пробелом в конце, как в эталоне
        reports["GamersBase"].to_excel(writer, sheet_name="продажи_GB ", index=False)
    return buf.getvalue()


def _fmt_money(v: float) -> str:
    return f"{v:,.2f}".replace(",", " ").replace(".", ",")


def _save_to_temp(uploaded) -> str:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(uploaded.getbuffer())
        return f.name


def _load_catalog_safe(uploaded):
    if uploaded is None:
        return None
    try:
        p = _save_to_temp(uploaded)
        cat = load_catalog(p)
        Path(p).unlink(missing_ok=True)
        return cat
    except Exception as e:
        st.warning(f"Не удалось прочитать каталог из {uploaded.name}: {e}")
        return None


def _load_partner_catalog_safe(uploaded):
    if uploaded is None:
        return None
    try:
        p = _save_to_temp(uploaded)
        pc = load_partner_catalog(p)
        Path(p).unlink(missing_ok=True)
        return pc
    except Exception as e:
        st.warning(f"Каталог партнёров не прочитан из {uploaded.name}: {e}")
        return None


# ---------- UI ---------------------------------------------------------------

st.set_page_config(page_title="SalesFlow", layout="wide")

st.markdown("# SalesFlow")
st.markdown("Сборка отчётов о продажах **B2B / Chinaplay / GamersBase** из Universal Report.")

uploaded = st.file_uploader(
    "Universal Report (xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
    help="Стандартная выгрузка биллинга `UniversalReport_DD_MM_YYYY_DD_MM_YYYY_shipped.xlsx`",
)

with st.expander("Дополнительно: каталоги для маппинга ID (необязательно)"):
    st.markdown(
        "Загрузите файлы-эталоны движения ключей (с листом `Каталог`), "
        "чтобы заменить ID биллинга на каталоговые ID и привести названия "
        "к виду «Новое название» (с региональным маркером)."
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        cat_b2b_file = st.file_uploader("B2B каталог", type=["xlsx"], key="catb2b",
                                         help="Файл `*_Движение_ключей_b2b_*.xlsx`")
    with col2:
        cat_chin_file = st.file_uploader("Chinaplay каталог", type=["xlsx"], key="catchin",
                                          help="Файл `*_Движение_ключей_Chinaplay_*.xlsx`")
    with col3:
        cat_gb_file = st.file_uploader("GamersBase каталог", type=["xlsx"], key="catgb",
                                        help="Файл `*_Движение_ключей_GB_*.xlsx`")

if not uploaded:
    st.info("Загрузите Universal Report, чтобы начать.")
    st.stop()


tmp_path = _save_to_temp(uploaded)
catalogs = {
    "B2B": _load_catalog_safe(cat_b2b_file),
    "Chinaplay": _load_catalog_safe(cat_chin_file),
    "GamersBase": _load_catalog_safe(cat_gb_file),
}
# Каталог партнёров — из того же B2B-файла (лист 'каталог партнеров')
partner_cat = _load_partner_catalog_safe(cat_b2b_file)

with st.spinner("Читаю и обрабатываю Universal Report..."):
    reports = build_all(tmp_path, catalogs=catalogs, partner_catalog=partner_cat)

Path(tmp_path).unlink(missing_ok=True)


# ----- Сводка ----------------------------------------------------------------

st.markdown("### Сводка")
s = summary(reports)
cols = st.columns(3)
for col, (_, row) in zip(cols, s.iterrows()):
    with col:
        st.metric(
            label=row["Отчёт"],
            value=f"{int(row['Кол-во ключей']):,} ключей".replace(",", " "),
            delta=f"{int(row['Строк'])} строк / {_fmt_money(row['Сумма выручки'])}",
            delta_color="off",
        )

# ----- Статистика маппинга ---------------------------------------------------

if any(c is not None for c in catalogs.values()):
    st.markdown("### Статус маппинга ID")
    mcols = st.columns(3)
    for col, name in zip(mcols, ["B2B", "Chinaplay", "GamersBase"]):
        with col:
            df = reports[name]
            stats = df["Маппинг"].value_counts()
            total = len(df)
            ok = int(stats.get("✓", 0)) + int(stats.get("✓ норм.", 0))
            pct = ok * 100 // total if total else 0
            st.markdown(f"**{name}**: {ok}/{total} ({pct}%)")
            st.dataframe(
                stats.reset_index().rename(columns={stats.name: "Строк", "index": "Статус"}),
                hide_index=True,
                use_container_width=True,
            )

st.divider()


# ----- Превью каждого отчёта -------------------------------------------------

tab_b2b, tab_chin, tab_gb = st.tabs(["Продажи Б2Б", "продажи_Чайна", "продажи_GB"])

with tab_b2b:
    df = reports["B2B"]
    st.markdown(
        f"**{len(df)} строк**  •  "
        f"{int(df['Количество проданных игр'].sum()):,} ключей".replace(",", " ")
    )
    by_curr = df.groupby("Валюта продажи", dropna=False).agg(
        rows=("ID", "count"),
        qty=("Количество проданных игр", "sum"),
        revenue=("Выручка от продажи", "sum"),
    ).reset_index()
    st.markdown("**По валютам:**")
    st.dataframe(by_curr, use_container_width=True, hide_index=True)
    st.markdown("**Детали:**")
    st.dataframe(df, use_container_width=True, hide_index=True, height=400)

with tab_chin:
    df = reports["Chinaplay"]
    st.markdown(
        f"**{len(df)} строк**  •  {int(df['Кол-во'].sum()):,} ключей  •  "
        f"{_fmt_money(df['Сумма'].sum())} CNY".replace(",", " ")
    )
    st.dataframe(df, use_container_width=True, hide_index=True)

with tab_gb:
    df = reports["GamersBase"]
    st.markdown(
        f"**{len(df)} строк**  •  {int(df['Кол-во'].sum()):,} ключей  •  "
        f"{_fmt_money(df['Сумма'].sum())} RUB".replace(",", " ")
    )
    st.dataframe(df, use_container_width=True, hide_index=True)


# ----- Download -------------------------------------------------------------

st.divider()
xlsx_bytes = _to_xlsx_bytes(reports)
st.download_button(
    label="Скачать xlsx со всеми тремя отчётами",
    data=xlsx_bytes,
    file_name=f"Продажи_свод_{uploaded.name.replace('.xlsx', '')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)


# ----- Технические заметки --------------------------------------------------

with st.expander("Логика сборки и известные ограничения"):
    st.markdown("""
**Фильтры площадок:**
- B2B → `площадка == 'продажи б2б'`
- Chinaplay → `площадка == 'Продажи Чайна'`
- GamersBase → `площадка == 'Продажи ГБ'`

**Колонки сумм:**
- B2B: `Сумма позиции заказа в валюте партнера`
- Chinaplay: `Сумма позиции заказа без учета комиссии` (внутренняя CNY-сумма)
- GamersBase: `Сумма позиции заказа без учета комиссии` (RUB)

**Партнёр:**
- B2B: `Имя партнера из платформы` если заполнено, иначе `Партнёр`
- Chinaplay: фикс. `Физическое лицо Chinaplay`
- GamersBase: фикс. `Физическое лицо Gamersbase`

**Маппинг ID (если загружен каталог):**
1. Точное совпадение названия (case-insensitive) → каталоговый ID и «Новое название»
2. Fallback: нормализованное сравнение (без ™®©, знаков препинания и регион-постфиксов RUB/USD/EUR/UAH/PLN/...)
3. Несколько кандидатов → ID не меняется, статус «⚠ N exact/норм.»
4. Не найдено → ID не меняется, статус «✗ не найден»

**Известные ограничения:**
1. **Суммы Chinaplay** расходятся с эталоном на ~2% из-за усреднения курсов CNY/RUB.
2. **Ручные продажи стока** (исторические продажи через другие подразделения) НЕ собираются —
   только то, что есть в биллинге за месяц.
3. **Маппинг по имени** не работает для переименованных продуктов и для случаев,
   когда в каталоге несколько регионов одного товара (требует ручного выбора).
""")
