"""
Plati-бандлы: один листинг продаёт несколько изданий (standard/deluxe/…),
поля издания в выгрузке нет, код ключа `товар` непрозрачен, справочника
код→ID нет. Решение (без справочника):

  1) base (игра) — по имени листинга через токен-матч к catalog_master;
  2) издание — кластеризация цен ВНУТРИ валюты оплаты по чистому разрыву;
     при чистом разрыве auto-split по тирам, иначе весь спрос на базовое
     издание + флаг 'издание неясно (промо)';
  3) регион — снаружи, через allocator (как для всех каналов).

Деньги/итог листинга не затрагиваются (они уже точны у билдера) — здесь только
раскладка СПРОСА по (base, издание) для движения.
"""
from __future__ import annotations
import re
import numpy as np
import pandas as pd

_WORD = re.compile(r"[a-zа-я0-9]+", re.I)
_TIER = [("ultimate", ["ultimate", "ультимейт", "ультиматив"]),
         ("complete", ["complete", "комплит", "полное"]),
         ("gold", ["gold", "голд"]),
         ("premium", ["premium", "премиум"]),
         ("deluxe", ["deluxe", "делюкс", "делюx"]),
         ("standard", ["standard", "standart", "стандарт", "базов"])]


def _tokens(s: str) -> set:
    return set(_WORD.findall(str(s).lower()))


def _tier_of(name: str) -> str:
    n = str(name).lower()
    for tier, kws in _TIER:
        if any(k in n for k in kws):
            return tier
    return "standard"


def _tier_rank(tier: str) -> int:
    order = {"standard": 0, "deluxe": 1, "premium": 2, "gold": 3,
             "complete": 4, "ultimate": 5}
    return order.get(tier, 0)


def _game_root(base: str) -> str:
    """Корень игры без ключевых слов издания (для группировки изданий базы)."""
    s = str(base)
    for _, kws in _TIER:
        for k in kws:
            s = re.sub(rf"[:\-–]?\s*\b{k}\b\s*(edition|издани\w*)?", "", s, flags=re.I)
    return re.sub(r"[:\-–]\s*$", "", re.sub(r"\s+", " ", s)).strip()


def catalog_editions(catalog: pd.DataFrame) -> dict:
    """{game_root: [base_name, ...]} — издания игры, упорядоченные от младшего.

    Порядок: сначала базовое (имя == корню игры или самое короткое = standard),
    затем по рангу ключевого слова тира и длине имени. Это позволяет маппить
    ценовые кластеры на издания по РАНГУ, не перечисляя имена спецредакций
    (Voidfarer, Super Citizen и т.п.).
    """
    groups: dict = {}
    for b in catalog["base"].dropna().unique():
        groups.setdefault(_game_root(b), []).append(str(b))
    out: dict = {}
    for root, bases in groups.items():
        def _rank(b):
            return (_tier_rank(_tier_of(b)), len(str(b)))
        out[root] = sorted(set(bases), key=_rank)
    return out


def match_root(listing: str, roots: dict) -> str | None:
    core = _strip_decoration(listing)
    core = re.sub(r"(выбор издания|choose the edition|std\b|deluxe|standart|"
                  r"ultimate|complete|premium|gold|edition|издани\w*|/.*)", "",
                  core, flags=re.I)
    toks = _tokens(core)
    if not toks:
        return None
    best, score = None, 0
    for r in roots:
        rt = _tokens(r)
        s = len(toks & rt)
        if s > score:
            score, best = s, r
    if best is None:
        return None
    if score >= 2 or (len(toks) <= 2 and toks <= _tokens(best)):
        return best
    return None


def _strip_decoration(listing: str) -> str:
    """Убирает зонный префикс/суффикс Plati (сегменты через '|') и хвосты."""
    s = str(listing)
    if "|" in s:
        parts = [p.strip() for p in s.split("|")]
        # выкидываем сегменты-зоны и 'STEAM КЛЮЧ/KEY'
        cand = [p for p in parts if not re.search(
            r"^(рф|снг|cis|ru|eu|tr|no ru|без рф|steam|key|ключ|💸|psn|playstation)",
            p, re.I) and len(p) > 2]
        s = max(cand, key=len) if cand else parts[len(parts) // 2]
    return s


def match_base(listing: str, cat_bases: dict) -> str | None:
    """Лучшее совпадение имени листинга с base по пересечению токенов."""
    core = _strip_decoration(listing)
    core = re.sub(r"(выбор издания|choose the edition|std\b|deluxe|standart|"
                  r"ultimate|complete|premium|gold|edition|издани\w*|/.*)", "",
                  core, flags=re.I)
    toks = _tokens(core)
    if not toks:
        return None
    best, score = None, 0
    for b in cat_bases:
        bt = _tokens(b)
        s = len(toks & bt)
        if s > score:
            score, best = s, b
    # порог: >=2 общих токена ИЛИ (короткое ядро и все его токены в базе)
    if best is None:
        return None
    if score >= 2:
        return best
    if len(toks) <= 2 and toks <= _tokens(best):
        return best
    return None


def _clean_split(prices: np.ndarray) -> tuple[float, bool]:
    """Граница между тирами по максимальному разрыву. bool = разрыв чистый."""
    p = np.unique(prices)
    if len(p) < 2:
        return (p[0] if len(p) else 0, True)
    gaps = np.diff(p)
    bi = int(np.argmax(gaps))
    mx = gaps[bi]
    second = float(np.sort(gaps)[-2]) if len(gaps) > 1 else 0.0
    med = float(np.median(p))
    clean = mx > 0.2 * med and (second == 0 or mx > 2.5 * second)
    return (p[bi + 1], clean)


def resolve(unmapped_listings: list, raw: pd.DataFrame, catalog: pd.DataFrame,
            name_col: str = "название товара", price_col: str = "оплачено",
            cur_col: str = "валюта") -> tuple[dict, pd.DataFrame]:
    """Возврат: (demand, review_df).

    demand: {catalog_base: qty} — спрос по каталожному base издания (ключ
    allocator: Название без региона). Издание определяется ценой; регион —
    снаружи, allocator.
    review: листинги, где игра не найдена или издание неясно (промо).
    """
    ce = catalog_editions(catalog)          # {game_root: [base_name, ...]}
    demand: dict = {}
    review = []

    def _add(base_name, qty):
        demand[base_name] = demand.get(base_name, 0) + int(qty)

    for listing in unmapped_listings:
        rows = raw[raw[name_col].astype(str) == str(listing)]
        if len(rows) == 0:
            review.append({"Листинг": listing, "Флаг": "нет строк в выгрузке"})
            continue
        root = match_root(listing, ce)
        if root is None:
            review.append({"Листинг": listing, "Кол": int(len(rows)),
                           "Флаг": "игра не найдена в каталоге"})
            continue
        editions = ce[root]                  # список base от младшего к старшему
        if len(editions) <= 1:
            _add(editions[0] if editions else root, len(rows))
            continue
        e_lo, e_hi = editions[0], editions[1]
        clean_all = True
        for cur, g in rows.groupby(cur_col):
            _, clean = _clean_split(np.asarray(g[price_col].values, float))
            if not clean:
                clean_all = False
                break
        for cur, g in rows.groupby(cur_col):
            pr = np.asarray(g[price_col].values, float)
            split, _ = _clean_split(np.unique(pr))
            _add(e_lo, int((pr < split).sum()))
            _add(e_hi, int((pr >= split).sum()))
        if not clean_all:
            review.append({"Листинг": listing, "Кол": int(len(rows)), "База": root,
                           "Флаг": "издание по промо-цене (приближённо)"})
    return demand, pd.DataFrame(review)


EXCLUDE_KEYS = ["автопополнение", "выгодно", "карта пополнения", "replenishment",
                "karta podarunkowa", "карта подарочная", "gift card", "itunes",
                "пополнение steam"]
# Номиналы карт и спецредакции — точный персистентный маппинг, не эвристика:
# heuristic путает 100↔500 PLN. Такие листинги → в ревью «новые листинги».


def suggest_from_raw(raw_path: str, mapping, catalog: pd.DataFrame,
                     name_col: str = "Name") -> tuple[dict, pd.DataFrame]:
    """Полный проход подсказок по бандлам из сырой выгрузки Plati.

    Использует чистую колонку `Name` (без зонной декорации) — ту же, по которой
    матчит билдер, чтобы НЕ дублировать уже сопоставленные листинги.
    Возврат: (demand {catalog_base: qty}, review_df).
    """
    raw = pd.read_excel(raw_path, sheet_name="выгрузка", engine="calamine")
    if name_col not in raw.columns:
        name_col = "название товара"
    norm = lambda s: re.sub(r"\s+", " ", str(s)).strip().lower()
    mapped = set(pd.Series(mapping["listing"]).map(norm)) if len(mapping) else set()
    raw = raw[raw[name_col].map(lambda s: not any(k in str(s).lower()
                                                  for k in EXCLUDE_KEYS))]
    raw["_k"] = raw[name_col].map(norm)
    un = raw[~raw["_k"].isin(mapped)]
    names = sorted(un[name_col].astype(str).unique())
    demand, review = resolve(names, raw, catalog, name_col=name_col)
    return demand, review
