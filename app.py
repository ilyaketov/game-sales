"""
SalesFlow — единое приложение (авто-эталон).

Вход БЕЗ эталона: стоячий catalog_master + персистентные маппинги (mappings/) +
биллинг (R1/R2/genba) + 6 выгрузок витрин + перенос остатков по (канал, ID) +
События (перемещения, опционально). Выход = собранное «Движение ключей»
(= авто-эталон) + QB-листы продаж. Новые листинги (нет в маппинге) → ревью
с подсказкой по имени из catalog_master.
"""
from __future__ import annotations

import io
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import orchestrator as orch
import extract_mappings as em
import plati_bundles as pb

st.set_page_config(page_title="SalesFlow — авто-эталон", layout="wide")

HERE = Path(__file__).parent
MAPDIR = HERE / "mappings"
INDIV = ["Eneba", "Kinguin", "Driffle", "G2A"]
_WORD = re.compile(r"[a-zа-я0-9]+", re.I)


def _save(u) -> str | None:
    if u is None:
        return None
    d = Path(tempfile.mkdtemp()); p = d / u.name
    p.write_bytes(u.getbuffer()); return str(p)


def _money(v) -> str:
    try: return f"{float(v):,.2f}".replace(",", " ")
    except Exception: return str(v)


def _load_mapping(ch: str):
    f = MAPDIR / f"{ch.lower()}_mapping.csv"
    if not f.exists():
        return None
    return em.load_ggsel_mapping(str(f)) if ch == "GGSel" else em.load_df_mapping(str(f))


def _opt_col(uploaded, col: str):
    """Опциональный файл [ID, <col>] (xlsx/csv)."""
    if uploaded is None:
        return None
    p = _save(uploaded)
    df = pd.read_csv(p) if p.endswith(".csv") else pd.read_excel(p, engine="calamine")
    df = df[pd.to_numeric(df["ID"], errors="coerce").notna()].copy()
    df["ID"] = df["ID"].astype("Int64")
    c = col if col in df.columns else df.columns[1]
    df[col] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df[["ID", col]]


def _load_carryover(uploaded) -> dict:
    """Перенос остатков → {канал: DataFrame[ID, Остаток_нач]}.

    Поддерживает per-channel файл [Канал, ID, Остаток_нач|Остаток_конец]
    (выход прошлого месяца) и простой [ID, Остаток_нач] (общий '*').
    """
    if uploaded is None:
        return {}
    p = _save(uploaded)
    df = pd.read_csv(p) if p.endswith(".csv") else pd.read_excel(p, engine="calamine")
    df = df[pd.to_numeric(df["ID"], errors="coerce").notna()].copy()
    df["ID"] = df["ID"].astype("Int64")
    valcol = next((c for c in ("Остаток_нач", "Остаток_конец") if c in df.columns),
                  df.columns[-1])
    df["Остаток_нач"] = pd.to_numeric(df[valcol], errors="coerce").fillna(0).astype(int)
    if "Канал" in df.columns:
        return {ch: g[["ID", "Остаток_нач"]].reset_index(drop=True)
                for ch, g in df.groupby("Канал")}
    return {"*": df[["ID", "Остаток_нач"]]}


def _carry(carry_in: dict, channel: str):
    return carry_in.get(channel, carry_in.get("*"))


def _suggest(name: str, cat: pd.DataFrame, cat_tokens: dict) -> str:
    """Подсказка catID по имени листинга (пересечение токенов с catalog_master)."""
    toks = set(_WORD.findall(str(name).lower()))
    if not toks:
        return ""
    best, score = None, 0
    for cid, ct in cat_tokens.items():
        s = len(toks & ct)
        if s > score:
            score, best = s, cid
    if best is None or score < 2:
        return ""
    nm = cat.loc[cat["ID"] == best, "Название"]
    return f"{best} — {nm.iloc[0]}" if len(nm) else str(best)


def _xlsx(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        for n, df in sheets.items():
            df.to_excel(w, sheet_name=n[:31], index=False)
    return buf.getvalue()


# ─────────────── UI ───────────────
st.title("SalesFlow — авто-эталон")
st.caption("Вход: биллинг + выгрузки + стоячие справочники (без эталона). "
           "Выход: собранное «Движение ключей» + QB-листы продаж.")

with st.sidebar:
    st.header("Биллинг (KeyFlow)")
    r1_f = st.file_uploader("R1 — Универсальный отчёт", type=["xlsx"], key="r1")
    r2_f = st.file_uploader("R2 — shipped", type=["xlsx"], key="r2")
    gen_f = st.file_uploader("genbaFile", type=["xlsx"], key="gen")
    st.divider()
    st.header("Справочники")
    cat_up = st.file_uploader("catalog_master (csv) — по умолч. встроенный",
                              type=["csv"], key="cat")
    st.caption(f"Маппинги: встроенные из mappings/ ({len(list(MAPDIR.glob('*_mapping.csv')))} файлов)")
    st.divider()
    st.header("Перенос / корректировки")
    nach_f = st.file_uploader("Остаток_нач (перенос конца прошлого мес.)",
                              type=["xlsx", "csv"], key="nach")
    ev_f = st.file_uploader("События (перемещения, опц.)", type=["xlsx", "csv"], key="ev")
    report_date = st.date_input("Дата отчёта", value=pd.Timestamp("2026-05-31"))

st.subheader("Выгрузки кабинетов")
raws = {}
cols = st.columns(3)
for i, ch in enumerate(INDIV + ["Plati", "GGSel"]):
    with cols[i % 3]:
        raws[ch] = st.file_uploader(f"{ch}", type=["xlsx"], key=f"raw_{ch}")

run = st.button("Собрать авто-эталон", type="primary")

if run:
    if not (r1_f and r2_f and gen_f):
        st.error("Нужны все три биллинговых файла."); st.stop()
    cat_path = _save(cat_up) if cat_up else str(MAPDIR / "catalog_master.csv")
    cat = orch.load_catalog_master(cat_path)
    cat_tokens = {int(r.ID): set(_WORD.findall(str(r.Название).lower()))
                  for r in cat.itertuples()}
    carry_in = _load_carryover(nach_f)
    ev = _opt_col(ev_f, "События")

    with st.spinner("Загрузка биллинга (KeyFlow)… ~40с"):
        pipe = orch.load_pipeline(_save(r1_f), _save(r2_f), _save(gen_f))
    rd = pd.Timestamp(report_date)
    results = {}
    raw_paths = {}

    for ch in INDIV:
        if raws[ch] is None:
            continue
        mp = _load_mapping(ch)
        rp = _save(raws[ch]); raw_paths[ch] = rp
        try:
            out = orch.run_channel_m(pipe, ch, rp, mp, cat,
                                     ostatok_nach=_carry(carry_in, ch), sobytiya=ev, report_date=rd)
            results[ch] = out
        except Exception as e:
            st.warning(f"{ch}: {type(e).__name__}: {e}")

    if raws["Plati"] and raws["GGSel"]:
        try:
            plati_raw = _save(raws["Plati"])
            raw_paths["Plati"] = plati_raw
            pm = _load_mapping("Plati")
            try:
                bundle_demand, bundles = pb.suggest_from_raw(plati_raw, pm, cat)
            except Exception:
                bundle_demand, bundles = {}, pd.DataFrame()
            out = orch.run_combined_m(
                pipe, ["Plati", "GGSel"],
                {"Plati": plati_raw, "GGSel": _save(raws["GGSel"])},
                {"Plati": pm, "GGSel": _load_mapping("GGSel")},
                cat, ostatok_nach=_carry(carry_in, "Plati+GGSel"), sobytiya=ev,
                zone="Plati", report_date=rd, extra_demand=bundle_demand)
            results["Plati+GGSel"] = {"sales_multi": out["sales"],
                                      "movement": out["movement"],
                                      "review": out["review"],
                                      "new_listings": out["new_listings"],
                                      "bundles": bundles}
        except Exception as e:
            st.warning(f"Plati+GGSel: {type(e).__name__}: {e}")

    # подсказки для новых листингов
    for ch, out in results.items():
        nl = out.get("new_listings")
        if nl is not None and len(nl):
            nl = nl.copy()
            nl["Подсказка (catalog_master)"] = nl["Наименование"].map(
                lambda n: _suggest(n, cat, cat_tokens))
            out["new_listings"] = nl

    st.session_state["res"] = results
    st.session_state["raw_paths"] = raw_paths

results = st.session_state.get("res")
if results:
    st.subheader("Сводка")
    rows = []
    for ch, out in results.items():
        items = out["sales_multi"].items() if "sales_multi" in out else [(ch, out["sales"])]
        for name, s in items:
            rows.append({"Канал": name, "Продаж (шт)": int(s["Количество"].sum()),
                         "Сумма": _money(s["Сумма"].sum()),
                         "Валюта": s["Валюта"].mode().iloc[0] if len(s) else "",
                         "Новых листингов": int(s["ID"].isna().sum())})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # новые листинги
    new_all = []
    for ch, out in results.items():
        nl = out.get("new_listings")
        if nl is not None and len(nl):
            n = nl.copy()
            if "Канал" not in n.columns:
                n.insert(0, "Канал", ch)
            new_all.append(n)
    new_all = pd.concat(new_all, ignore_index=True) if new_all else pd.DataFrame()
    st.subheader(f"Новые листинги (нет в маппинге) — {len(new_all)}")
    if len(new_all):
        st.caption("Подтвердите catID (предзаполнен подсказкой), затем соберите "
                   "обновлённые маппинги для коммита в mappings/. "
                   "Каналы, ключённые по listing: Eneba/Driffle/G2A/Plati.")

        def _sugg_id(s):
            m = re.match(r"\s*(\d+)", str(s))
            return int(m.group(1)) if m else None

        ed = new_all.copy()
        ed["catID (подтвердить)"] = ed.get("Подсказка (catalog_master)", "").map(_sugg_id)
        showc = [c for c in ["Канал", "Наименование", "Количество", "Сумма",
                             "Подсказка (catalog_master)", "catID (подтвердить)"]
                 if c in ed.columns]
        edited = st.data_editor(
            ed[showc], use_container_width=True, hide_index=True, height=280,
            column_config={"catID (подтвердить)": st.column_config.NumberColumn(
                "catID (подтвердить)", help="ID из catalog_master; пусто = оставить в ревью")},
            disabled=[c for c in showc if c != "catID (подтвердить)"], key="ed_new")

        if st.button("Собрать обновлённые маппинги (для mappings/)"):
            conf = edited[edited["catID (подтвердить)"].notna()].copy()
            raw_paths = st.session_state.get("raw_paths", {})
            updated = {}
            for ch, g in conf.groupby("Канал"):
                add = pd.DataFrame({"listing": g["Наименование"].astype(str),
                                    "catID": g["catID (подтвердить)"],
                                    "product": g["Наименование"].astype(str)})
                try:
                    if ch == "Kinguin":
                        cur = _load_mapping("Kinguin")
                        rp = raw_paths.get("Kinguin")
                        if rp is None:
                            continue
                        updated["kinguin"] = em.merge_kinguin_mapping(cur, add, rp)
                    elif ch == "GGSel":
                        cur = pd.read_csv(MAPDIR / "ggsel_mapping.csv")
                        updated["ggsel"] = em.merge_ggsel_mapping(cur, add)
                    else:  # Eneba/Driffle/G2A/Plati/Plati+GGSel — по listing
                        tgt = "plati" if ch == "Plati+GGSel" else ch.lower()
                        cur = _load_mapping("Plati" if ch == "Plati+GGSel" else ch)
                        cur = cur if isinstance(cur, pd.DataFrame) else pd.DataFrame(
                            columns=["listing", "catID", "product"])
                        updated[tgt] = em.merge_listing_mapping(cur, add)
                except Exception as e:
                    st.warning(f"{ch}: не удалось слить — {type(e).__name__}: {e}")
            if not updated:
                st.warning("Нет подтверждённых строк с catID.")
            else:
                import zipfile
                zbuf = io.BytesIO()
                with zipfile.ZipFile(zbuf, "w") as z:
                    for tgt, df in updated.items():
                        z.writestr(f"{tgt}_mapping.csv", df.to_csv(index=False))
                st.success("Обновлены: " + ", ".join(f"{t} ({len(updated[t])} строк)"
                                                      for t in updated))
                st.download_button("Скачать обновлённые *_mapping.csv (zip)",
                                   data=zbuf.getvalue(),
                                   file_name="mappings_updated.zip",
                                   mime="application/zip")
    else:
        st.success("Все листинги сопоставлены.")

    # Plati-бандлы: подсказки game+издание (несопоставленные листинги)
    pg = results.get("Plati+GGSel", {})
    bundles = pg.get("bundles")
    if bundles is not None and len(bundles):
        st.subheader(f"Plati-бандлы на подтверждение — {len(bundles)}")
        st.caption("Игра/издание распознаны эвристикой (имя + цена). Подтвердите и "
                   "добавьте в mappings/plati_mapping.csv (персистентно).")
        st.dataframe(bundles, use_container_width=True, hide_index=True, height=200)

    # ревью аллокатора
    rev_all = []
    for ch, out in results.items():
        r = out["review"]
        if len(r):
            r = r.copy(); r.insert(0, "Канал", ch); rev_all.append(r)
    rev_all = pd.concat(rev_all, ignore_index=True) if rev_all else pd.DataFrame()
    st.subheader(f"Ревью разнесения — {len(rev_all)}")
    if len(rev_all):
        st.dataframe(rev_all, use_container_width=True, hide_index=True, height=240)

    tabs = st.tabs(list(results.keys()))
    for tab, (ch, out) in zip(tabs, results.items()):
        with tab:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Продажи (QB)**")
                if "sales_multi" in out:
                    for name, s in out["sales_multi"].items():
                        st.caption(name)
                        st.dataframe(s, use_container_width=True, hide_index=True, height=200)
                else:
                    st.dataframe(out["sales"], use_container_width=True,
                                 hide_index=True, height=360)
            with c2:
                st.markdown("**Движение ключей (авто-эталон)**")
                st.dataframe(out["movement"], use_container_width=True,
                             hide_index=True, height=360)

    st.divider()
    sheets = {}
    for ch, out in results.items():
        if "sales_multi" in out:
            for name, s in out["sales_multi"].items():
                sheets[f"Продажи_{name}"] = s
            sheets["Движение_Plati_GGSel"] = out["movement"]
        else:
            sheets[f"Продажи_{ch}"] = out["sales"]
            sheets[f"Движение_{ch}"] = out["movement"]
    if len(new_all):
        sheets["Новые_листинги"] = new_all
    if len(rev_all):
        sheets["Ревью"] = rev_all
    carry = []
    for ch, out in results.items():
        m = out["movement"][["ID", "Название", "Регион", "Остаток_конец"]].copy()
        m = m.rename(columns={"Остаток_конец": "Остаток_нач"})  # round-trip: вход след. месяца
        m.insert(0, "Канал", ch); carry.append(m)
    if carry:
        sheets["Перенос_остатков"] = pd.concat(carry, ignore_index=True)
    st.download_button("Скачать авто-эталон (xlsx)", data=_xlsx(sheets),
                       file_name=f"SalesFlow_{pd.Timestamp(report_date):%Y_%m}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
