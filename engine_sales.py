"""
SalesFlow engine.

Принимает Universal Report за месяц, возвращает три отчёта-сводки в форматах:
- "Продажи Б2Б"
- "продажи_Чайна"
- "продажи_GB "  (с пробелом!)

Логика:
- B2B (площадка == 'продажи б2б'):
    * Партнёр = Имя партнера из платформы (если не пусто) иначе Партнёр
    * Валюта = Валюта партнера
    * Сумма = SUM(Сумма позиции заказа в валюте партнера)
    * Кол-во = SUM(Количество)
    * Цена = Сумма / Кол-во
    * Группировка: ID, Продукт, Партнёр_итог, Валюта

- Chinaplay (площадка == 'Продажи Чайна'):
    * Партнёр = "Физическое лицо Chinaplay" (константа)
    * Валюта = CNY (константа)
    * Сумма = SUM(Сумма позиции заказа без учета комиссии)
    * Кол-во = SUM(Количество)
    * Группировка: ID, Продукт

- GamersBase (площадка == 'Продажи ГБ'):
    * Партнёр = "Физическое лицо Gamersbase" (константа)
    * Валюта = RUB
    * Сумма = SUM(Сумма позиции заказа без учета комиссии)
    * Кол-во = SUM(Количество)
    * Группировка: ID, Продукт
"""
from __future__ import annotations

import pandas as pd
from pathlib import Path
from typing import Dict, Optional


# Колонки каталога эталона
CAT_ID = "ID"
CAT_PRODUCT = "Продукт"
CAT_NEW_NAME = "Новое назавание"  # Опечатка эталона — сохраняем

# Колонки из universal report
COL_PLOSHADKA = "площадка"
COL_ID = "ID продукта"
COL_PRODUCT = "Продукт"
COL_PARTNER = "Партнёр"
COL_PARTNER_PLATFORM = "Имя партнера из платформы"
COL_QTY = "Количество"
COL_SUM_PARTNER = "Сумма позиции заказа в валюте партнера"
COL_SUM_NO_COMM = "Сумма позиции заказа без учета комиссии"
COL_CURRENCY = "Валюта партнера"
COL_SUPPLIER = "Поставщик"

# Фильтры площадок
PL_B2B = "продажи б2б"
PL_CHINAPLAY = "Продажи Чайна"
PL_GB = "Продажи ГБ"

# Фиксированные значения
CONST_PARTNER_CHINAPLAY = "Физическое лицо Chinaplay"
CONST_PARTNER_GB = "Физическое лицо Gamersbase"
CONST_CURRENCY_CHINAPLAY = "CNY"
CONST_CURRENCY_GB = "RUB"


def _load_universal(path: str | Path) -> pd.DataFrame:
    """Читает Universal Report (xlsx). Берёт первый лист, начинающийся на UniversalReport."""
    xls = pd.ExcelFile(path)
    target = None
    for s in xls.sheet_names:
        if s.startswith("UniversalReport"):
            target = s
            break
    if target is None:
        # Fallback: лист с наибольшим количеством строк
        target = max(xls.sheet_names, key=lambda s: xls.parse(s, nrows=1).shape[1])
    df = pd.read_excel(path, sheet_name=target, engine="calamine")
    return df


def extract_period(path: str | Path) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Извлекает (start_date, end_date) из имени файла UniversalReport.

    Формат имени: 'UniversalReport_DD_MM_YYYY_DD_MM_YYYY_shipped.xlsx'.
    Если не удаётся — возвращает (NaT, NaT).
    """
    import re
    name = Path(path).name
    m = re.search(r"(\d{1,2})_(\d{1,2})_(\d{4})_(\d{1,2})_(\d{1,2})_(\d{4})", name)
    if not m:
        return pd.NaT, pd.NaT
    d1, m1, y1, d2, m2, y2 = map(int, m.groups())
    try:
        return pd.Timestamp(year=y1, month=m1, day=d1), pd.Timestamp(year=y2, month=m2, day=d2)
    except ValueError:
        return pd.NaT, pd.NaT


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


# ----- Каталог партнёров (B2B) ----------------------------------------------

def load_partner_catalog(path: str | Path) -> pd.DataFrame:
    """Читает лист 'каталог партнеров' из xlsx-файла эталона B2B,
    дополняет его записями из листа 'Партнеры ROKKY ' (если есть) и
    из EXTRA_PARTNER_CATALOG.

    Возвращает DataFrame: ЮЛ, Партнёр_биллинг, Название_1С.
    """
    df = pd.read_excel(path, sheet_name="каталог партнеров", header=None,
                       engine="calamine")
    if df.shape[1] < 3:
        raise ValueError("Лист 'каталог партнеров' должен иметь ≥3 колонок")
    out = df.iloc[:, :3].copy()
    out.columns = ["ЮЛ", "Партнёр_биллинг", "Название_1С"]
    out = out.dropna(subset=["Партнёр_биллинг"])
    out = out[out["Партнёр_биллинг"] != "Партнер по биллингу"].reset_index(drop=True)
    out["ЮЛ"] = out["ЮЛ"].astype(str).str.strip().replace({"ТМ": "TM"})

    # Дополнение из листа 'Партнеры ROKKY ' (если есть)
    out = _load_rokky_sheet(path, out)
    # Дополнение из встроенного справочника
    out = _merge_with_extras(out)

    # Ключи поиска
    out["_key"] = out["Партнёр_биллинг"].astype(str).str.strip().str.lower()
    out["_norm_key"] = out["Партнёр_биллинг"].apply(_normalize_partner)
    return out


def _load_rokky_sheet(path: str | Path, catalog: pd.DataFrame) -> pd.DataFrame:
    """Подтягивает партнёров из листа 'Партнеры ROKKY ' (если такой есть),
    добавляя их в каталог с ЮЛ=ROKKY.
    """
    try:
        rk = pd.read_excel(path, sheet_name="Партнеры ROKKY ", header=None,
                           engine="calamine")
    except Exception:
        return catalog
    if rk.empty or rk.shape[1] < 1:
        return catalog
    # Первая колонка — имя партнёра, остальное не важно
    names = rk.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    existing = set(catalog["Партнёр_биллинг"].astype(str).str.strip().str.lower())
    extras = []
    for name in names:
        if name and name.lower() not in existing:
            extras.append({"ЮЛ": "ROKKY", "Партнёр_биллинг": name,
                           "Название_1С": name})
            existing.add(name.lower())
    if not extras:
        return catalog
    return pd.concat([catalog, pd.DataFrame(extras)], ignore_index=True)


# Список партнёров, которые в каталоге значатся ЮЛ=ROKKY, но в эталоне
# Продажи Б2Б оставляются под своими именами (не маркируются как ROKKY).
# Это эмпирическое исключение, основанное на наблюдении за эталоном апреля 2026.
ROKKY_KEEP_NAME: set[str] = {
    "(EU) CiDiKi",   # 5 строк эталона под именем CiDiKi, не ROKKY
}


# Партнёры с ЮЛ=GE, для которых в эталоне используется их имя платформы
# (а не «Название_1С» из каталога). Эмпирическое наблюдение апреля 2026.
GE_KEEP_NAME: set[str] = {
    "Wei",   # 6 строк эталона под именем 'Wei', не Inner Mongolia Qidian...
}


# Правообладатели продуктов, которые для прямых продаж (CIS) IGM
# переводят транзакцию в FZE-канал (вместо TM).
# Эмпирическое правило: KRM-каналы (PSN-карты, AppStore Gift Cards)
# проходят через FZE-юр.лицо.
IGM_FZE_RIGHTS_HOLDERS: set[str] = {"KRM"}


# Дополнительный встроенный справочник для случаев, когда нужный лист
# отсутствует в файле эталона. После пополнения каталога в эталоне может
# быть пустым.
EXTRA_PARTNER_CATALOG: list[tuple[str, str, str]] = [
    # FZE-партнёры платформы (подтверждено пользователем):
    # AlfaBank, GamersHub, MobiMoney идут под ЮЛ=FZE.
    # Игромир (= (CIS) IGM) → FZE для KRM-правообладателей продуктов
    # (обрабатывается отдельным правилом IGM_FZE_RIGHTS_HOLDERS).
    ("FZE", "AlfaBank",  "AlfaBank"),
    ("FZE", "GamersHub", "GamersHub"),
    ("FZE", "MobiMoney", "MobiMoney"),
    # ROKKY-партнёры, отсутствующие в листе 'Партнеры ROKKY ', но
    # маркированные ROKKY в эталоне:
    ("ROKKY", "(CIS) game-online", "(CIS) game-online"),
]


def _merge_with_extras(catalog: pd.DataFrame) -> pd.DataFrame:
    """Добавляет EXTRA_PARTNER_CATALOG к каталогу. Существующие записи
    с тем же 'Партнёр_биллинг' не перетираются."""
    if not EXTRA_PARTNER_CATALOG:
        return catalog
    existing = set(catalog["Партнёр_биллинг"].astype(str).str.strip().str.lower())
    extras = []
    for yul, pname, name1c in EXTRA_PARTNER_CATALOG:
        if pname.strip().lower() not in existing:
            extras.append({"ЮЛ": yul, "Партнёр_биллинг": pname, "Название_1С": name1c})
    if not extras:
        return catalog
    return pd.concat([catalog, pd.DataFrame(extras)], ignore_index=True)


def _normalize_partner(s) -> str:
    """Нормализация имени партнёра:
    - убираем региональные префиксы (CIS), (CN), (EU), (TRY), (RU), (TR), (CNT)
    - убираем доменные суффиксы .com/.ru/.io/.net/.org/.shop
    - сохраняем (VAT) — сильный признак российских партнёров (ТМ)
    - lowercase, нормализуем пробелы
    """
    import re
    if not isinstance(s, str):
        return ""
    x = s.strip()
    x = re.sub(r"\((CIS|CN|EU|TRY|RU|TR|CNT|EUR|USD|RUB|CNY|GBP)\)\s*",
               "", x, flags=re.IGNORECASE)
    x = re.sub(r"\.(com|ru|io|net|org|shop)$", "", x, flags=re.IGNORECASE)
    x = re.sub(r"\s+", " ", x).strip().lower()
    return x


def _map_b2b_partner(billing_partner: str, partner_platform: str | None,
                    p_cat_platform: Optional[dict],
                    p_cat_direct: Optional[dict],
                    product_rights_holder: Optional[str] = None) -> tuple[str, str | None]:
    """Возвращает (Партнёр_итог, Партнёр_на_платформе) для строки B2B.

    Логика:
    1. Если Имя_платформы заполнено → ищем его в p_cat_platform.
    2. Иначе → ищем Партнёр биллинга в p_cat_direct.
    3. По найденному ЮЛ:
        - TM/ROKKY/FZE → используем ЮЛ как имя партнёра
        - GE → используем Название_1С (длинное юр.имя)
    4. Если не найдено в каталоге:
        - Имя_платформы → оставляем имя платформы
        - Иначе → оставляем Партнёр биллинга
    5. Эмпирические исключения:
        - ROKKY_KEEP_NAME (например, (EU) CiDiKi) → оставляем имя
        - GE_KEEP_NAME (например, Wei) → оставляем имя платформы
        - (CIS) IGM + правообладатель в IGM_FZE_RIGHTS_HOLDERS → FZE
    """
    has_platform = (isinstance(partner_platform, str)
                    and partner_platform.strip()
                    and partner_platform.strip().lower() != "nan")
    platform_name = partner_platform.strip() if has_platform else None

    # Спецправило: (CIS) IGM как прямой партнёр + KRM-правообладатель → FZE
    if (not has_platform
            and isinstance(billing_partner, str)
            and billing_partner.strip() == "(CIS) IGM"
            and isinstance(product_rights_holder, str)
            and product_rights_holder.strip() in IGM_FZE_RIGHTS_HOLDERS):
        return "FZE", None

    if not p_cat_platform and not p_cat_direct:
        return (platform_name or str(billing_partner)), platform_name

    # Выбираем словарь и кандидата для поиска
    if has_platform:
        cand = platform_name
        p_cat = p_cat_platform
    else:
        cand = billing_partner
        p_cat = p_cat_direct

    if not p_cat or not isinstance(cand, str) or not cand.strip():
        return (platform_name or str(billing_partner)), platform_name

    # exact match → fallback на normalized
    rec = p_cat.get(cand.strip().lower())
    if rec is None:
        rec = p_cat.get(_normalize_partner(cand))

    if rec is None:
        return (platform_name or str(billing_partner)), platform_name

    found_yul = rec["ЮЛ"].strip() if isinstance(rec["ЮЛ"], str) else ""
    found_name1c = rec["Название_1С"] if isinstance(rec["Название_1С"], str) else None

    if found_yul in ("TM", "ROKKY", "FZE"):
        if found_yul == "ROKKY" and isinstance(cand, str) and cand.strip() in ROKKY_KEEP_NAME:
            return cand.strip(), platform_name
        return found_yul, platform_name
    if found_yul == "GE":
        # Wei и подобные — остаются под именем платформы
        if isinstance(cand, str) and cand.strip() in GE_KEEP_NAME:
            return cand.strip(), platform_name
        if found_name1c and found_name1c.strip():
            return found_name1c.strip(), platform_name

    return (platform_name or str(billing_partner)), platform_name


# ----- Каталог: загрузка и маппинг ------------------------------------------

def load_catalog(path: str | Path) -> pd.DataFrame:
    """Читает лист 'Каталог' из xlsx-файла эталона.

    Возвращает DataFrame с колонками: ID, Продукт, Новое название, Правообладатель.
    Опечатка 'назавание' (вместо 'название') — оригинальная из эталона B2B.
    """
    df = pd.read_excel(path, sheet_name="Каталог", engine="calamine")
    cols = {c.strip(): c for c in df.columns if isinstance(c, str)}
    id_col = cols.get(CAT_ID) or list(df.columns)[0]
    prod_col = cols.get(CAT_PRODUCT) or list(df.columns)[1]
    name_col = cols.get(CAT_NEW_NAME)
    if name_col is None:
        for c in df.columns:
            if isinstance(c, str) and "новое" in c.lower():
                name_col = c
                break
    # Правообладатель — опциональная колонка
    rh_col = None
    for c in df.columns:
        if isinstance(c, str) and "правообладатель" in c.lower():
            rh_col = c
            break

    keep = [c for c in [id_col, prod_col, name_col, rh_col] if c is not None]
    target_cols = ["ID", "Продукт", "Новое название", "Правообладатель"][: len(keep)]
    out = df[keep].copy()
    out.columns = target_cols
    out = out.dropna(subset=["ID", "Продукт"]).reset_index(drop=True)
    out["ID"] = pd.to_numeric(out["ID"], errors="coerce").astype("Int64")
    out["Продукт_lower"] = out["Продукт"].astype(str).str.strip().str.lower()
    return out


def _normalize_name(s: str) -> str:
    """Нормализация для fuzzy-маппинга:
    - lowercase, strip
    - убираем ™ ® ©
    - убираем знаки препинания , . : ; - / | _
    - убираем регион-постфиксы (RUB, USD, EUR, UAH, PLN, KRW, CNY, NOK, GBP)
    - нормализуем пробелы
    """
    import re
    if not isinstance(s, str):
        return ""
    x = s.strip().lower()
    x = re.sub(r"[™®©]", "", x)
    # Убираем регион-постфиксы (может быть несколько: "Game RUB EUR")
    region_re = r"\b(rub|usd|eur|uah|pln|krw|cny|nok|gbp|jpy)\b"
    prev = None
    while prev != x:
        prev = x
        x = re.sub(rf"{region_re}\s*$", "", x).strip()
    # Унифицируем разделители
    x = re.sub(r"[,.\:;\-/|_]+", " ", x)
    x = re.sub(r"\s+", " ", x).strip()
    return x


def _build_name_index(catalog: pd.DataFrame) -> dict[str, list[dict]]:
    """Индекс: lowercase product name → список записей каталога."""
    idx: dict[str, list[dict]] = {}
    for _, r in catalog.iterrows():
        key = r["Продукт_lower"]
        rec = {"ID": r["ID"], "Новое название": r.get("Новое название", r["Продукт"])}
        idx.setdefault(key, []).append(rec)
    return idx


def _build_normalized_index(catalog: pd.DataFrame) -> dict[str, list[dict]]:
    idx: dict[str, list[dict]] = {}
    for _, r in catalog.iterrows():
        key = _normalize_name(str(r["Продукт"]))
        if not key:
            continue
        rec = {"ID": r["ID"], "Новое название": r.get("Новое название", r["Продукт"])}
        idx.setdefault(key, []).append(rec)
    return idx


def _lookup(product_raw: str, exact_idx: dict, norm_idx: dict,
            billing_id: Optional[int] = None,
            id_idx: Optional[dict] = None) -> tuple[Optional[int], Optional[str], str]:
    """Возвращает (id, new_name, status).

    Приоритет:
    1. Если billing_id есть в каталоге → используем его (имя берём из каталога)
    2. Exact match по названию
    3. Normalized match
    """
    # 1) Биллинговый ID уже в каталоге
    if billing_id is not None and id_idx is not None:
        try:
            rec = id_idx.get(int(billing_id))
        except (TypeError, ValueError):
            rec = None
        if rec is not None:
            return int(rec["ID"]), (rec.get("Новое название") or product_raw), "✓ по ID"

    key = str(product_raw).strip().lower()
    matches = exact_idx.get(key, [])
    if len(matches) == 1:
        return int(matches[0]["ID"]), matches[0]["Новое название"], "✓"
    if len(matches) > 1:
        return None, None, f"⚠ {len(matches)} exact"

    nkey = _normalize_name(product_raw)
    nmatches = norm_idx.get(nkey, []) if nkey else []
    if len(nmatches) == 1:
        return int(nmatches[0]["ID"]), nmatches[0]["Новое название"], "✓ норм."
    if len(nmatches) > 1:
        return None, None, f"⚠ {len(nmatches)} норм."
    return None, None, "✗ не найден"


def _build_id_index(catalog: pd.DataFrame) -> dict:
    """Индекс ID → запись каталога (для приоритетного поиска по биллинговому ID)."""
    idx = {}
    for _, r in catalog.iterrows():
        try:
            idx[int(r["ID"])] = r
        except (TypeError, ValueError):
            pass
    return idx


def _apply_catalog(df: pd.DataFrame, catalog: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Заменяет ID и Наименование на каталоговые. Колонка 'Маппинг' не добавляется
    (исключена для QuickBooks-совместимого вывода).
    """
    out = df.copy()
    if catalog is None or out.empty:
        return out

    exact = _build_name_index(catalog)
    norm = _build_normalized_index(catalog)
    id_idx = _build_id_index(catalog)

    new_ids: list = []
    new_names: list = []
    for _, r in out.iterrows():
        billing_id = r["ID"] if pd.notna(r["ID"]) else None
        cid, cname, _st = _lookup(r["Наименование"], exact, norm, billing_id, id_idx)
        if cid is not None:
            new_ids.append(cid)
            new_names.append(cname or r["Наименование"])
        else:
            new_ids.append(r["ID"])
            new_names.append(r["Наименование"])

    out["ID"] = pd.array(new_ids, dtype="Int64")
    out["Наименование"] = new_names
    return out


def _apply_catalog_b2b(df: pd.DataFrame, catalog: Optional[pd.DataFrame]) -> pd.DataFrame:
    """Для B2B аналогично: колонка 'Наименование' (унифицированная)."""
    out = df.copy()
    if catalog is None or out.empty:
        return out

    exact = _build_name_index(catalog)
    norm = _build_normalized_index(catalog)
    id_idx = _build_id_index(catalog)

    new_ids: list = []
    new_names: list = []
    for _, r in out.iterrows():
        billing_id = r["ID"] if pd.notna(r["ID"]) else None
        cid, cname, _st = _lookup(r["Наименование"], exact, norm, billing_id, id_idx)
        if cid is not None:
            new_ids.append(cid)
            new_names.append(cname or r["Наименование"])
        else:
            new_ids.append(r["ID"])
            new_names.append(r["Наименование"])

    out["ID"] = pd.array(new_ids, dtype="Int64")
    out["Наименование"] = new_names
    return out


def build_b2b(df: pd.DataFrame,
              partner_catalog: Optional[pd.DataFrame] = None,
              product_catalog: Optional[pd.DataFrame] = None,
              report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """Сборка отчёта B2B в формате 'Продажи Б2Б'.

    Из биллинга исключаются подзаказы Chinaplay и GamersBase
    (Имя партнера из платформы ∈ {Chinaplay, GamersBase}).

    product_catalog — опциональный каталог продуктов с колонкой 'Правообладатель'.
                      Используется для правила (CIS) IGM → FZE при KRM.
    report_date — дата отчёта (последний день периода). Подставляется в колонку 'Дата'.
    """
    sub = df[df[COL_PLOSHADKA] == PL_B2B].copy()
    if sub.empty:
        return _empty_b2b()

    # Исключаем чужие подзаказы — они принадлежат другим площадкам
    sub = sub[~sub[COL_PARTNER_PLATFORM].isin(["Chinaplay", "GamersBase"])].copy()
    if sub.empty:
        return _empty_b2b()

    # Индекс правообладателей по ID продукта (для правила (CIS) IGM → FZE)
    rights_holder_idx: dict = {}
    if product_catalog is not None and not product_catalog.empty:
        for _, r in product_catalog.iterrows():
            try:
                pid = int(r["ID"])
            except (TypeError, ValueError):
                continue
            # Колонка с правообладателем может быть 'Правообладатель' или 'Rights_holder'
            rh = None
            for col in ("Правообладатель", "Rights_holder", "rights_holder"):
                if col in r and isinstance(r[col], str) and r[col].strip():
                    rh = r[col].strip()
                    break
            if rh:
                rights_holder_idx[pid] = rh

    # Подготовим два словаря каталога партнёров: для платформенных имён и для
    # прямых партнёров. Для одного партнёра в каталоге может быть несколько ЮЛ
    # (например, (CIS) IGM = TM в одном контексте, ROKKY в другом). Правило
    # эвристики:
    #   - Поиск по Имя_партнёра_из_платформы → приоритет ROKKY (платформенные продажи)
    #   - Поиск по Партнёр биллинга (без платформы) → приоритет TM (прямые с НДС)
    p_cat_platform: Optional[dict] = None
    p_cat_direct: Optional[dict] = None
    if partner_catalog is not None and not partner_catalog.empty:
        # Сортируем так, чтобы при коллизии нужный ЮЛ перетёр другие
        cat_for_platform = partner_catalog.copy()
        cat_for_platform["_pri"] = cat_for_platform["ЮЛ"].map(
            {"ROKKY": 0, "FZE": 1, "GE": 2, "TM": 3}).fillna(9)
        cat_for_platform = cat_for_platform.sort_values("_pri", ascending=False)
        p_cat_platform = {}
        for _, r in cat_for_platform.iterrows():
            for col in ("_key", "_norm_key"):
                if col in r and isinstance(r[col], str) and r[col]:
                    p_cat_platform[r[col]] = r

        cat_for_direct = partner_catalog.copy()
        cat_for_direct["_pri"] = cat_for_direct["ЮЛ"].map(
            {"TM": 0, "FZE": 1, "GE": 2, "ROKKY": 3}).fillna(9)
        cat_for_direct = cat_for_direct.sort_values("_pri", ascending=False)
        p_cat_direct = {}
        for _, r in cat_for_direct.iterrows():
            for col in ("_key", "_norm_key"):
                if col in r and isinstance(r[col], str) and r[col]:
                    p_cat_direct[r[col]] = r

    # Маппинг партнёров: (Партнёр_итог, Партнёр_на_платформе)
    def _resolve_rh(pid):
        try:
            return rights_holder_idx.get(int(pid))
        except (TypeError, ValueError):
            return None

    mapped = sub.apply(
        lambda r: _map_b2b_partner(r[COL_PARTNER], r[COL_PARTNER_PLATFORM],
                                    p_cat_platform, p_cat_direct,
                                    _resolve_rh(r[COL_ID])),
        axis=1, result_type="expand",
    )
    sub["_partner"] = mapped[0]
    sub["_qty"] = _to_numeric(sub[COL_QTY])
    sub["_revenue"] = _to_numeric(sub[COL_SUM_PARTNER])

    agg = sub.groupby([COL_ID, COL_PRODUCT, "_partner", COL_CURRENCY],
                      dropna=False, as_index=False).agg(
        qty=("_qty", "sum"),
        revenue=("_revenue", "sum"),
    )
    agg["price"] = (agg["revenue"] / agg["qty"]).where(agg["qty"] != 0, 0)

    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": agg[COL_ID].astype("Int64"),
        "Наименование": agg[COL_PRODUCT].astype("string"),
        "Партнер": agg["_partner"].astype("string"),
        "Количество": agg["qty"].astype("Int64"),
        "Цена": pd.to_numeric(agg["price"], errors="coerce").round(6),
        "Валюта": agg[COL_CURRENCY].astype("string"),
        "Сумма": pd.to_numeric(agg["revenue"], errors="coerce").round(2),
    })
    out = out[UNIFIED_COLUMNS]
    out = out.sort_values(["ID", "Партнер"], na_position="last").reset_index(drop=True)
    return out


def _empty_b2b() -> pd.DataFrame:
    return pd.DataFrame(columns=UNIFIED_COLUMNS)


UNIFIED_COLUMNS = ["Дата", "ID", "Наименование", "Партнер", "Количество", "Цена", "Валюта", "Сумма"]


def build_chinaplay(df: pd.DataFrame,
                     report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """Сборка отчёта Chinaplay в формате 'продажи_Чайна'."""
    sub = df[df[COL_PLOSHADKA] == PL_CHINAPLAY].copy()
    if sub.empty:
        return _empty_simple()

    sub["_qty"] = _to_numeric(sub[COL_QTY])
    sub["_revenue"] = _to_numeric(sub[COL_SUM_NO_COMM])

    agg = sub.groupby([COL_ID, COL_PRODUCT],
                      dropna=False, as_index=False).agg(
        qty=("_qty", "sum"),
        revenue=("_revenue", "sum"),
    )
    agg["price"] = (agg["revenue"] / agg["qty"]).where(agg["qty"] != 0, 0)

    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": agg[COL_ID].astype("Int64"),
        "Наименование": agg[COL_PRODUCT].astype("string"),
        "Партнер": pd.Series([CONST_PARTNER_CHINAPLAY] * len(agg), dtype="string"),
        "Количество": agg["qty"].astype("Int64"),
        "Цена": pd.to_numeric(agg["price"], errors="coerce").round(6),
        "Валюта": pd.Series([CONST_CURRENCY_CHINAPLAY] * len(agg), dtype="string"),
        "Сумма": pd.to_numeric(agg["revenue"], errors="coerce").round(2),
    })
    out = out[UNIFIED_COLUMNS]
    return out.sort_values("ID").reset_index(drop=True)


def build_gb(df: pd.DataFrame,
             report_date: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """Сборка отчёта GamersBase в формате 'продажи_GB '."""
    sub = df[df[COL_PLOSHADKA] == PL_GB].copy()
    if sub.empty:
        return _empty_simple()

    sub["_qty"] = _to_numeric(sub[COL_QTY])
    sub["_revenue"] = _to_numeric(sub[COL_SUM_NO_COMM])

    agg = sub.groupby([COL_ID, COL_PRODUCT],
                      dropna=False, as_index=False).agg(
        qty=("_qty", "sum"),
        revenue=("_revenue", "sum"),
    )
    agg["price"] = (agg["revenue"] / agg["qty"]).where(agg["qty"] != 0, 0)

    out = pd.DataFrame({
        "Дата": pd.to_datetime(report_date) if report_date is not None else pd.NaT,
        "ID": agg[COL_ID].astype("Int64"),
        "Наименование": agg[COL_PRODUCT].astype("string"),
        "Партнер": pd.Series([CONST_PARTNER_GB] * len(agg), dtype="string"),
        "Количество": agg["qty"].astype("Int64"),
        "Цена": pd.to_numeric(agg["price"], errors="coerce").round(6),
        "Валюта": pd.Series([CONST_CURRENCY_GB] * len(agg), dtype="string"),
        "Сумма": pd.to_numeric(agg["revenue"], errors="coerce").round(2),
    })
    out = out[UNIFIED_COLUMNS]
    return out.sort_values("ID").reset_index(drop=True)


def _empty_simple() -> pd.DataFrame:
    return pd.DataFrame(columns=UNIFIED_COLUMNS)


def build_all(
    path: str | Path,
    catalogs: Optional[Dict[str, pd.DataFrame]] = None,
    partner_catalog: Optional[pd.DataFrame] = None,
    report_date: Optional[pd.Timestamp] = None,
) -> Dict[str, pd.DataFrame]:
    """Возвращает все три отчёта по одному пути universal report.

    catalogs — опциональный словарь {'B2B': df, 'Chinaplay': df, 'GamersBase': df}
               где значения — DataFrame'ы каталогов продуктов (load_catalog).
    partner_catalog — опциональный DataFrame каталога партнёров B2B.
    report_date — дата отчёта для колонки 'Дата'. Если не указана —
                  автоматически определяется как конец периода из имени файла.
    """
    df = _load_universal(path)
    if report_date is None:
        _start, end = extract_period(path)
        report_date = end if pd.notna(end) else None

    catalogs = catalogs or {}
    return {
        "B2B": _apply_catalog_b2b(
            build_b2b(df, partner_catalog=partner_catalog,
                       product_catalog=catalogs.get("B2B"),
                       report_date=report_date),
            catalogs.get("B2B")),
        "Chinaplay": _apply_catalog(
            build_chinaplay(df, report_date=report_date),
            catalogs.get("Chinaplay")),
        "GamersBase": _apply_catalog(
            build_gb(df, report_date=report_date),
            catalogs.get("GamersBase")),
    }


def summary(reports: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Краткая сводка для UI."""
    rows = []
    for name, dfr in reports.items():
        rows.append({
            "Отчёт": name,
            "Строк": len(dfr),
            "Кол-во ключей": int(pd.to_numeric(dfr["Количество"], errors="coerce").sum())
                              if len(dfr) else 0,
            "Сумма выручки": float(pd.to_numeric(dfr["Сумма"], errors="coerce").sum())
                              if len(dfr) else 0,
        })
    return pd.DataFrame(rows)
