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
import report_xlsx
import storage

st.set_page_config(page_title="SalesFlow — прогнать всё", layout="wide")

try:
    _GH = storage.cfg_from_secrets(st.secrets)
except Exception:
    _GH = None

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


@st.cache_resource(show_spinner=False)
def _cached_pipe(r1: str, r2: str, genba: str):
    """Загруженный Pipeline кэшируется по хэш-путям биллинга: пока файлы те же,
    повторные взаимодействия не перезагружают ~50с биллинг."""
    return orch.load_pipeline(r1, r2, genba)


def _money(v) -> str:
    try:
        return f"{float(v):,.2f}".replace(",", " ")
    except Exception:
        return str(v)


def _map_file(name: str) -> str | None:
    """Ищет справочник и в mappings/, и рядом с app.py (на случай плоской выкладки)."""
    for base in (MAPDIR, HERE):
        p = base / name
        if p.exists():
            return str(p)
    return None


def _load_mapping(ch: str):
    f = _map_file(f"{ch.lower()}_mapping.csv")
    if f is None:
        return None
    return em.load_ggsel_mapping(f) if ch == "GGSel" else em.load_df_mapping(f)


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


def _load_events(path: str):
    """События: {канал: df} если есть колонка Канал, иначе общий df [ID, События]."""
    if path is None:
        return None
    df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path, engine="calamine")
    df = df[pd.to_numeric(df["ID"], errors="coerce").notna()].copy()
    df["ID"] = df["ID"].astype("Int64")
    col = "События" if "События" in df.columns else df.columns[-1]
    df["События"] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    if "Канал" in df.columns:
        return {ch: g[["ID", "События"]].reset_index(drop=True)
                for ch, g in df.groupby("Канал")}
    return df[["ID", "События"]]


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
    with st.expander("Дополнительно (каталог · GitHub)"):
        cat_up = st.file_uploader("catalog_master.csv (по умолч. встроенный)",
                                  type=["csv"], key="cat")
        if storage.available(_GH):
            ok, msg = storage.check(_GH)
            (st.success if ok else st.warning)(f"GitHub: {msg}")
        else:
            st.caption("GitHub не настроен — маппинги локальные (mappings/). Для "
                       "автосохранения добавьте [github] token/repo в секреты (см. README).")

if not up:
    st.info("Загрузите хотя бы биллинг (R1/R2/genba) и выгрузки кабинетов.")
    st.stop()

# распознаём (кэш по имени+размеру, чтобы не перечитывать на каждом rerun)
sig = tuple((f.name, f.size) for f in up)
if st.session_state.get("_detect_sig") != sig:
    det, unknown, conflicts = {}, [], []
    with st.spinner("Распознаём файлы…"):
        for f in up:
            p = _save_upload(f)
            k = detect.detect_kind(p)
            if k is None:
                unknown.append(f.name)
            elif k in det:
                conflicts.append((k, det[k]["name"], f.name))   # дубль типа
                det[k] = {"name": f.name, "path": p, "size": f.size}
            else:
                det[k] = {"name": f.name, "path": p, "size": f.size}
    st.session_state["_detect_sig"] = sig
    st.session_state["_detected"] = det
    st.session_state["_unknown"] = unknown
    st.session_state["_conflicts"] = conflicts
det = st.session_state["_detected"]
unknown = st.session_state["_unknown"]
conflicts = st.session_state.get("_conflicts", [])

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
if conflicts:
    for k, old, new in conflicts:
        st.warning(f"Два файла распознались как «{KIND_LABEL[k]}»: `{old}` и `{new}` — "
                   f"взят последний (`{new}`). Проверьте, не лишний ли файл.")

missing_bill = [k for k in ("r1", "r2", "genba") if k not in det]
raw_channels = [k for k in ["Eneba", "Kinguin", "Driffle", "G2A", "Plati", "GGSel"] if k in det]

# доступные юниты продаж из распознанного (Plati+GGSel — только если оба файла)
SALES_UNITS = ["Eneba", "Kinguin", "Driffle", "G2A", "Plati+GGSel"]
avail_units = [u for u in SALES_UNITS
               if (u in det) or (u == "Plati+GGSel" and "Plati" in det and "GGSel" in det)]

st.divider()
mode = st.radio("Режим прогона", ["Прогнать всё", "Только закуп", "Только продажи"],
                horizontal=True)

sel_zones = orch.PURCHASE_ZONES
sel_units = avail_units
if mode == "Только закуп":
    sel_zones = st.multiselect("Площадки закупа", orch.PURCHASE_ZONES,
                               default=orch.PURCHASE_ZONES)
elif mode == "Только продажи":
    sel_units = st.multiselect("Каналы продаж", SALES_UNITS, default=avail_units,
                               help="Plati+GGSel считается единым юнитом (нужны оба файла)")

# требования под режим
need_bill = bool(missing_bill)          # биллинг нужен всегда (Загружено из pipe)
need_units = (mode != "Только закуп")
problems = []
if need_bill:
    problems.append("биллинг: " + ", ".join(KIND_LABEL[m] for m in missing_bill))
if need_units and not sel_units:
    problems.append("не выбран ни один канал продаж (или не загружены их выгрузки)")

BTN = {"Прогнать всё": "Прогнать всё (закуп + продажи)",
       "Только закуп": "Прогнать закуп",
       "Только продажи": "Прогнать продажи"}
run = st.button(BTN[mode], type="primary", disabled=bool(problems))
if problems:
    st.error("Не хватает для запуска — " + "; ".join(problems))


# ─────────────── ЗАПУСК: закуп / продажи / всё ───────────────
def _apply_suggestions(sales_results, cat, cat_tokens):
    for ch, out in sales_results.items():
        nl = out.get("new_listings")
        if nl is not None and len(nl):
            nl = nl.copy()
            nl["Подсказка (catalog_master)"] = nl["Наименование"].map(
                lambda n: _suggest(n, cat, cat_tokens))
            out["new_listings"] = nl


def _zones_for_units(units):
    z = []
    for u in units:
        if u == "Plati+GGSel":
            z.append("Plati")
        else:
            z.append(u)
    return z


if run:
    need_cat = (mode != "Только закуп")
    cat, cat_tokens = None, {}
    if need_cat:
        cat_path = _save_upload(cat_up) if cat_up else _map_file("catalog_master.csv")
        if cat_path is None:
            st.error("Не найден catalog_master.csv (ни в mappings/, ни рядом с app.py). "
                     "Загрузите его в «Дополнительно (каталог)» или добавьте в репозиторий "
                     "в папку mappings/.")
            st.stop()
        try:
            cat = orch.load_catalog_master(cat_path)
        except Exception as e:
            st.error(f"Не удалось прочитать catalog_master.csv: {type(e).__name__}: {e}")
            st.stop()
        cat_tokens = {int(r.ID): set(_WORD.findall(str(r.Название).lower()))
                      for r in cat.itertuples()}
    carry_in = _load_carryover(det["carry"]["path"]) if "carry" in det else {}
    events = _load_events(det["events"]["path"]) if "events" in det else None
    rd = pd.Timestamp(report_date)

    with st.spinner("Загрузка биллинга (~50с; кэшируется до смены файлов)…"):
        pipe = _cached_pipe(det["r1"]["path"], det["r2"]["path"], det["genba"]["path"])

    # какие raws/mappings задействуем (по выбранным юнитам продаж)
    def _raws_for(units):
        r = {}
        if "Plati+GGSel" in units and "Plati" in det and "GGSel" in det:
            r["Plati"] = det["Plati"]["path"]; r["GGSel"] = det["GGSel"]["path"]
        for u in units:
            if u in INDIV and u in det:
                r[u] = det[u]["path"]
        return r

    raws = _raws_for(sel_units) if mode != "Только закуп" else {}
    mappings = {ch: _load_mapping(ch) for ch in raws}

    # Plati-бандлы (если Plati в игре)
    bundles = pd.DataFrame(); bundle_demand = {}
    if "Plati" in raws:
        try:
            bundle_demand, bundles = pb.suggest_from_raw(raws["Plati"], mappings["Plati"], cat)
        except Exception as e:
            st.warning(f"Plati-бандлы: {type(e).__name__}: {e}")

    res = {"mode": mode, "errors": [], "events_fed": events is not None}
    if mode == "Только закуп":
        with st.spinner("Прогон закупа…"):
            pur = orch.run_purchases(pipe, zones=sel_zones)
        res["purchases"] = pur
        res["purchases_flat"] = orch.purchases_flat(pur)

    elif mode == "Только продажи":
        with st.spinner("Прогон продаж…"):
            sales = orch.run_sales_all(pipe, cat, raws, mappings, carry=carry_in,
                                       events=events, report_date=rd,
                                       extra_demand=bundle_demand, errors=res["errors"])
            pur = orch.run_purchases(pipe, zones=_zones_for_units(sel_units))  # для гейта
        _apply_suggestions(sales, cat, cat_tokens)
        if "Plati+GGSel" in sales:
            sales["Plati+GGSel"]["bundles"] = bundles
        res["sales"] = sales
        res["reconcile"] = orch.reconcile(pur, sales)
        res["review"] = orch.review_consolidated(sales, catalog=cat)

    else:  # Прогнать всё
        with st.spinner("Прогон закупа и продаж…"):
            full = orch.run_everything(pipe, cat, raws, mappings, carry=carry_in,
                                       events=events, report_date=rd,
                                       extra_demand=bundle_demand)
        _apply_suggestions(full["sales"], cat, cat_tokens)
        if "Plati+GGSel" in full["sales"]:
            full["sales"]["Plati+GGSel"]["bundles"] = bundles
        res.update(full)

    st.session_state["res"] = res
    st.session_state["raws"] = raws
    st.session_state["rdate"] = pd.Timestamp(report_date)


# ─────────────── РЕНДЕР ───────────────
res = st.session_state.get("res")
if res:
    results = res.get("sales", {})
    rdate = st.session_state.get("rdate", pd.Timestamp(report_date))
    st.caption(f"Режим: **{res.get('mode', 'Прогнать всё')}**")

    # Ошибки каналов (изоляция сбоев)
    for ch, msg in res.get("errors", []):
        st.warning(f"Канал «{ch}» пропущен из-за ошибки: {msg}")

    # ── БЕЙДЖИ ВАЛИДАЦИИ ──
    b = st.columns(4)
    # гейт
    if "reconcile" in res:
        rec0 = res["reconcile"]
        bad0 = rec0[(rec0["Юнит"] != "—") & (rec0["Δ"] != 0)]
        b[0].metric("Гейт Δ", "0 ✅" if len(bad0) == 0 else f"{len(bad0)} ❌",
                    help="Загружено движения = закуп зоны")
    # закуп: неклассиф. поставщики
    if "purchases" in res:
        unc = sum(int(v["unclassified"]["Количество"].sum())
                  for v in res["purchases"].values() if len(v["unclassified"]))
        b[1].metric("Закуп: неклассиф.", "0 ✅" if unc == 0 else f"{unc} ⚠",
                    help="Ключи от нераспознанных поставщиков (возвращены в свод, подсвечены)")
    # склад / события
    if results:
        b[2].metric("Склад", "закрыт Событиями ✅" if res.get("events_fed")
                    else "без Событий ○",
                    help="С Событиями Остаток_конец сходится с учётом перемещений")
    # ручная сверка
    if "review" in res:
        b[3].metric("Ручная сверка", f"{len(res['review'])}",
                    help="Позиции для ручной проверки (регион/листинг)")

    # 1) ГЕЙТ (если есть сверка)
    if "reconcile" in res:
        st.subheader("Гейт сверки (Загружено движения ↔ закуп зоны)")
        rec = res["reconcile"]
        bad = rec[(rec["Юнит"] != "—") & (rec["Δ"] != 0)]
        if len(bad) == 0:
            st.success("Δ = 0 по всем юнитам движения — закуп и «Загружено» сходятся.")
        else:
            st.error(f"Расхождение в {len(bad)} юнит(ах) — проверьте проводку.")
        st.dataframe(rec, use_container_width=True, hide_index=True)

    # 2) СВОДКА ПРОДАЖ (если есть продажи)
    if results:
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

    # 3) ЗАКУП (если есть)
    if "purchases" in res:
        st.subheader(f"Закуп — {len(res['purchases'])} площадок")
        prows = [{"Площадка": z, "Товаров": len(v["flat"]), "Кол-во": v["qty"],
                  "Поставщиков": v["suppliers"], "Себестоимость": _money(v["cost"])}
                 for z, v in res["purchases"].items()]
        st.dataframe(pd.DataFrame(prows), use_container_width=True, hide_index=True)

    # 4) НОВЫЕ ЛИСТИНГИ (editor + сбор маппингов) — только если есть продажи
    new_all = []
    for ch, out in results.items():
        nl = out.get("new_listings")
        if nl is not None and len(nl):
            n = nl.copy()
            if "Канал" not in n.columns:
                n.insert(0, "Канал", ch)
            new_all.append(n)
    new_all = pd.concat(new_all, ignore_index=True) if new_all else pd.DataFrame()
    if results:
        st.subheader(f"Новые листинги (нет в маппинге) — {len(new_all)}")
    if results and len(new_all):
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
                        _gg = _map_file("ggsel_mapping.csv")
                        cur_gg = pd.read_csv(_gg) if _gg else pd.DataFrame(columns=["listing", "catID", "product"])
                        updated["ggsel"] = em.merge_ggsel_mapping(cur_gg, add)
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
                if storage.available(_GH):
                    saved = storage.save_many(_GH, updated)
                    if saved:
                        st.success("Закоммичено в GitHub (" + storage.repo_label(_GH)
                                   + "): " + ", ".join(saved)
                                   + ". Streamlit передеплоит — маппинги подхватятся.")
                    else:
                        st.info("В GitHub не сохранено (проверьте токен/права репозитория).")
                st.download_button("Скачать обновлённые *_mapping.csv (zip)",
                                   data=zbuf.getvalue(), file_name="mappings_updated.zip",
                                   mime="application/zip")
    elif results:
        st.success("Все листинги сопоставлены.")

    # 5) Plati-бандлы
    pg = results.get("Plati+GGSel", {})
    bundles = pg.get("bundles")
    if bundles is not None and len(bundles):
        st.subheader(f"Plati-бандлы на подтверждение — {len(bundles)}")
        st.caption("Игра/издание распознаны эвристикой (имя + цена). Подтвердите и "
                   "добавьте в mappings/plati_mapping.csv.")
        st.dataframe(bundles, use_container_width=True, hide_index=True, height=200)

    # 6) РУЧНАЯ СВЕРКА (консолидированная) — если есть
    if "review" in res:
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

    # 7) ВКЛАДКИ по каналам — если есть продажи
    if results:
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

    # 8) СКАЧАТЬ аннотированный xlsx (цвет + комментарии по проблемам)
    st.divider()
    tag = {"Прогнать всё": "всё", "Только закуп": "закуп",
           "Только продажи": "продажи"}.get(res.get("mode"), "всё")
    st.caption("В файле проблемные ячейки подсвечены цветом и снабжены "
               "комментариями; лист «Легенда и проблемы» — расшифровка.")
    st.download_button(f"Скачать результат ({tag}, xlsx)",
                       data=report_xlsx.build(res),
                       file_name=f"SalesFlow_{tag}_{rdate:%Y_%m}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # №5: перенос остатков как отдельный CSV для следующего месяца.
    # Тот же файл в окне распознаётся как «carry» и подаётся как Остаток_нач.
    if results:
        cf = []
        for ch, out in results.items():
            m = out["movement"][["ID", "Остаток_конец"]].copy()
            m = m.rename(columns={"Остаток_конец": "Остаток_нач"})
            m.insert(0, "Канал", ch)
            cf.append(m)
        cf = pd.concat(cf, ignore_index=True)
        cf = cf[pd.to_numeric(cf["Остаток_нач"], errors="coerce").fillna(0) != 0]
        st.download_button(
            "Скачать перенос остатков (Остаток_нач для след. месяца)",
            data=cf.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"carry_forward_{rdate:%Y_%m}.csv", mime="text/csv",
            help="В следующем месяце загрузите этот файл в то же окно — он "
                 "распознается как Остаток_нач и подхватится автоматически.")
