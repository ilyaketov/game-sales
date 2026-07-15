"""
SalesFlow — единое приложение «Прогнать всё» (закуп + продажи).

ОДНО ОКНО ЗАГРУЗКИ: киньте разом биллинг (R1/R2/genba) + 6 выгрузок кабинетов
(Eneba/Kinguin/Driffle/G2A/Plati/GGSel) и — опционально — Остаток_нач и События.
Приложение само распознаёт роль каждого файла по сигнатуре колонок и одной
кнопкой прогоняет всё: закуп (9 площадок), продажи+движение (6 каналов), гейт
сверки (Загружено↔закуп) и единый лист ручной сверки.
"""
from __future__ import annotations

import io
import re
import tempfile
import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

import orchestrator as orch
import extract_mappings as em
import plati_bundles as pb
import config
import detect

st.set_page_config(page_title="SalesFlow — прогнать всё", layout="wide")

HERE = Path(__file__).parent
MAPDIR = HERE / "mappings"
INDIV = ["Eneba", "Kinguin", "Driffle", "G2A"]
_WORD = re.compile(r"[a-zа-я0-9]+", re.I)
_TMP = Path(tempfile.mkdtemp())

KIND_LABEL = detect.KIND_LABEL
CORE = detect.CORE


def _save_upload(u) -> str:
    data = u.getbuffer()
    digest = hashlib.sha1(bytes(data)).hexdigest()[:12]
    p = _TMP / f"{digest}_{u.name}"
    if not p.exists():
        p.write_bytes(data)
    return str(p)


def _money(v) -> str:
    try:
        return f"{float(v):,.2f}".replace(",", " ")
    except Exception:
        return str(v)


def _load_mapping(ch: str):
    f = MAPDIR / f"{ch.lower()}_mapping.csv"
    if not f.exists():
        return None
    return em.load_ggsel_mapping(str(f)) if ch == "GGSel" else em.load_df_mapping(str(f))


def _load_opt_col(path: str, col: str):
    if path is None:
        return None
    df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path, engine="calamine")
    df = df[pd.to_numeric(df["ID"], errors="coerce").notna()].copy()
    df["ID"] = df["ID"].astype("Int64")
    c = col if col in df.columns else df.columns[1]
    df[col] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    return df[["ID", col]]


def _load_carryover(path: str) -> dict:
    if path is None:
        return {}
    df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path, engine="calamine")
    df = df[pd.to_numeric(df["ID"], errors="coerce").notna()].copy()
    df["ID"] = df["ID"].astype("Int64")
    valcol = next((c for c in ("Остаток_нач", "Остаток_конец") if c in df.columns),
                  df.columns[-1])
    df["Остаток_нач"] = pd.to_numeric(df[valcol], errors="coerce").fillna(0).astype(int)
    if "Канал" in df.columns:
        return {ch: g[["ID", "Остаток_нач"]].reset_index(drop=True)
                for ch, g in df.groupby("Канал")}
    return {"*": df[["ID", "Остаток_нач"]]}


def _suggest(name: str, cat: pd.DataFrame, cat_tokens: dict) -> str:
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


# ─────────────── UI: одно окно загрузки ───────────────
st.title("SalesFlow — прогнать всё")
st.caption("Одно окно: киньте биллинг (R1/R2/genba) + 6 выгрузок кабинетов "
           "(+ опц. Остаток_нач, События). Роль каждого файла распознаётся сама.")

up = st.file_uploader("Перетащите сюда все файлы разом",
                      type=["xlsx", "csv"], accept_multiple_files=True,
                      key="all_files")

c1, c2 = st.columns([2, 3])
with c1:
    report_date = st.date_input("Дата отчёта", value=pd.Timestamp("2026-05-31"))
with c2:
    with st.expander("Дополнительно (каталог)"):
        cat_up = st.file_uploader("catalog_master.csv (по умолч. встроенный)",
                                  type=["csv"], key="cat")

if not up:
    st.info("Загрузите хотя бы биллинг (R1/R2/genba) и выгрузки кабинетов.")
    st.stop()

# распознаём (кэш по имени+размеру, чтобы не перечитывать на каждом rerun)
sig = tuple((f.name, f.size) for f in up)
if st.session_state.get("_detect_sig") != sig:
    det, unknown = {}, []
    with st.spinner("Распознаём файлы…"):
        for f in up:
            p = _save_upload(f)
            k = detect.detect_kind(p)
            if k is None:
                unknown.append(f.name)
            else:
                det[k] = {"name": f.name, "path": p, "size": f.size}
    st.session_state["_detect_sig"] = sig
    st.session_state["_detected"] = det
    st.session_state["_unknown"] = unknown
det = st.session_state["_detected"]
unknown = st.session_state["_unknown"]

# карточки распознанного
st.markdown("**Распознано:**")
grid = st.columns(3)
order = CORE + ["carry", "events"]
for i, k in enumerate(order):
    with grid[i % 3]:
        if k in det:
            st.success(f"✓ {KIND_LABEL[k]}\n\n`{det[k]['name']}`")
        elif k in CORE:
            opt = k in ("carry", "events")
            st.markdown(f"{'○' if opt else '—'} {KIND_LABEL[k]} · _не загружен_")
for k in ("carry", "events"):
    if k in det:
        with grid[(order.index(k)) % 3]:
            st.info(f"✓ {KIND_LABEL[k]} (опц.)\n\n`{det[k]['name']}`")
if unknown:
    st.warning("Не распознаны (пропущены): " + ", ".join(unknown))

missing = [k for k in ("r1", "r2", "genba") if k not in det]
raw_channels = [k for k in ["Eneba", "Kinguin", "Driffle", "G2A", "Plati", "GGSel"] if k in det]

run = st.button("Прогнать всё (закуп + продажи)", type="primary",
                disabled=bool(missing))
if missing:
    st.error("Для запуска нужен весь биллинг. Не хватает: "
             + ", ".join(KIND_LABEL[m] for m in missing))


# ─────────────── ЗАПУСК: единый run_everything ───────────────
if run:
    cat_path = _save_upload(cat_up) if cat_up else str(MAPDIR / "catalog_master.csv")
    cat = orch.load_catalog_master(cat_path)
    cat_tokens = {int(r.ID): set(_WORD.findall(str(r.Название).lower()))
                  for r in cat.itertuples()}
    carry_in = _load_carryover(det["carry"]["path"]) if "carry" in det else {}
    events = _load_opt_col(det["events"]["path"], "События") if "events" in det else None

    raws = {ch: det[ch]["path"] for ch in raw_channels}
    mappings = {ch: _load_mapping(ch) for ch in raw_channels}

    # Plati-бандлы → extra_demand (+ таблица на подтверждение)
    bundles = pd.DataFrame()
    bundle_demand = {}
    if "Plati" in raws:
        try:
            bundle_demand, bundles = pb.suggest_from_raw(raws["Plati"], mappings["Plati"], cat)
        except Exception as e:
            st.warning(f"Plati-бандлы: {type(e).__name__}: {e}")

    with st.spinner("Загрузка биллинга + прогон закупа и продаж… (~60с)"):
        pipe = orch.load_pipeline(det["r1"]["path"], det["r2"]["path"], det["genba"]["path"])
        res = orch.run_everything(
            pipe, cat, raws, mappings, carry=carry_in, events=events,
            report_date=pd.Timestamp(report_date), extra_demand=bundle_demand)

    # подсказки для новых листингов + бандлы к Plati+GGSel
    for ch, out in res["sales"].items():
        nl = out.get("new_listings")
        if nl is not None and len(nl):
            nl = nl.copy()
            nl["Подсказка (catalog_master)"] = nl["Наименование"].map(
                lambda n: _suggest(n, cat, cat_tokens))
            out["new_listings"] = nl
    if "Plati+GGSel" in res["sales"]:
        res["sales"]["Plati+GGSel"]["bundles"] = bundles

    st.session_state["res"] = res
    st.session_state["raws"] = raws
    st.session_state["rdate"] = pd.Timestamp(report_date)


# ─────────────── РЕНДЕР ───────────────
res = st.session_state.get("res")
if res:
    results = res["sales"]
    rdate = st.session_state.get("rdate", pd.Timestamp(report_date))

    # 1) ГЕЙТ
    st.subheader("Гейт сверки (Загружено движения ↔ закуп зоны)")
    rec = res["reconcile"]
    bad = rec[(rec["Юнит"] != "—") & (rec["Δ"] != 0)]
    if len(bad) == 0:
        st.success("Δ = 0 по всем юнитам движения — закуп и «Загружено» сходятся.")
    else:
        st.error(f"Расхождение в {len(bad)} юнит(ах) — проверьте проводку.")
    st.dataframe(rec, use_container_width=True, hide_index=True)

    # 2) СВОДКА ПРОДАЖ
    st.subheader("Продажи — сводка")
    rows = []
    for ch, out in results.items():
        items = out["sales_multi"].items() if "sales_multi" in out else [(ch, out["sales"])]
        for name, s in items:
            rows.append({"Канал": name, "Продаж (шт)": int(s["Количество"].sum()),
                         "Сумма": _money(s["Сумма"].sum()),
                         "Валюта": s["Валюта"].mode().iloc[0] if len(s) else "",
                         "Новых листингов": int(s["ID"].isna().sum())})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # 3) ЗАКУП
    st.subheader("Закуп — 9 площадок")
    prows = [{"Площадка": z, "Товаров": len(v["flat"]), "Кол-во": v["qty"],
              "Поставщиков": v["suppliers"], "Себестоимость": _money(v["cost"])}
             for z, v in res["purchases"].items()]
    st.dataframe(pd.DataFrame(prows), use_container_width=True, hide_index=True)

    # 4) НОВЫЕ ЛИСТИНГИ (editor + сбор маппингов)
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
        st.caption("Подтвердите catID (предзаполнен подсказкой) и соберите обновлённые "
                   "маппинги для коммита в mappings/.")

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
            raws_ss = st.session_state.get("raws", {})
            updated = {}
            for ch, g in conf.groupby("Канал"):
                add = pd.DataFrame({"listing": g["Наименование"].astype(str),
                                    "catID": g["catID (подтвердить)"],
                                    "product": g["Наименование"].astype(str)})
                try:
                    if ch == "Kinguin":
                        rp = raws_ss.get("Kinguin")
                        if rp is None:
                            continue
                        updated["kinguin"] = em.merge_kinguin_mapping(_load_mapping("Kinguin"), add, rp)
                    elif ch == "GGSel":
                        updated["ggsel"] = em.merge_ggsel_mapping(pd.read_csv(MAPDIR / "ggsel_mapping.csv"), add)
                    else:
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
                st.success("Обновлены: " + ", ".join(f"{t} ({len(updated[t])})" for t in updated))
                st.download_button("Скачать обновлённые *_mapping.csv (zip)",
                                   data=zbuf.getvalue(), file_name="mappings_updated.zip",
                                   mime="application/zip")
    else:
        st.success("Все листинги сопоставлены.")

    # 5) Plati-бандлы
    pg = results.get("Plati+GGSel", {})
    bundles = pg.get("bundles")
    if bundles is not None and len(bundles):
        st.subheader(f"Plati-бандлы на подтверждение — {len(bundles)}")
        st.caption("Игра/издание распознаны эвристикой (имя + цена). Подтвердите и "
                   "добавьте в mappings/plati_mapping.csv.")
        st.dataframe(bundles, use_container_width=True, hide_index=True, height=200)

    # 6) РУЧНАЯ СВЕРКА (консолидированная)
    review = res["review"]
    st.subheader(f"Ручная сверка — {len(review)}")
    if len(review):
        cat_sum = (review.groupby("Категория")
                   .agg(Строк=("Позиция", "size"), Нераспределено=("Нераспределено", "sum"))
                   .reset_index())
        st.dataframe(cat_sum, use_container_width=True, hide_index=True)
        st.dataframe(review, use_container_width=True, hide_index=True, height=280)
    else:
        st.success("Нечего сверять вручную.")

    # 7) ВКЛАДКИ по каналам
    st.subheader("По каналам: продажи (QB) и движение (авто-эталон)")
    tabs = st.tabs(list(results.keys()))
    for tab, (ch, out) in zip(tabs, results.items()):
        with tab:
            a, b = st.columns(2)
            with a:
                st.markdown("**Продажи (QB)**")
                if "sales_multi" in out:
                    for name, s in out["sales_multi"].items():
                        st.caption(name)
                        st.dataframe(s, use_container_width=True, hide_index=True, height=200)
                else:
                    st.dataframe(out["sales"], use_container_width=True, hide_index=True, height=360)
            with b:
                st.markdown("**Движение ключей (авто-эталон)**")
                st.dataframe(out["movement"], use_container_width=True, hide_index=True, height=360)

    # 8) СКАЧАТЬ единый xlsx
    st.divider()
    sheets = {"Гейт": rec, "Ручная сверка": review, "Закуп (свод)": res["purchases_flat"]}
    for z, v in res["purchases"].items():
        sheets[f"Закуп {z}"] = v["flat"]
    for ch, out in results.items():
        if "sales_multi" in out:
            for name, s in out["sales_multi"].items():
                sheets[f"Продажи {name}"] = s
            sheets["Движение Plati+GGSel"] = out["movement"]
        else:
            sheets[f"Продажи {ch}"] = out["sales"]
            sheets[f"Движение {ch}"] = out["movement"]
    if len(new_all):
        sheets["Новые листинги"] = new_all
    carry = []
    for ch, out in results.items():
        m = out["movement"][["ID", "Название", "Регион", "Остаток_конец"]].copy()
        m = m.rename(columns={"Остаток_конец": "Остаток_нач"})
        m.insert(0, "Канал", ch)
        carry.append(m)
    if carry:
        sheets["Перенос остатков"] = pd.concat(carry, ignore_index=True)
    st.download_button("Скачать всё (xlsx)", data=_xlsx(sheets),
                       file_name=f"SalesFlow_прогнать_всё_{rdate:%Y_%m}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
