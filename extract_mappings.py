"""
Извлечение стоячих справочников из эталонов (разовый seed).

Выход (папка mappings/):
  catalog_master.csv         — мастер-каталог: ID, Продукт, Регион, по,
                               Правообладатель, Новое назавание (объединение
                               по всем эталонам, dedupe по ID).
  kinguin_mapping.csv        — PRODUCTID, listing, catID, product
  driffle_mapping.csv        — listing, catID, product
  eneba_mapping.csv          — listing, catID, product
  g2a_mapping.csv            — listing, catID, product
  plati_mapping.csv          — listing, catID, product
  ggsel_mapping.csv          — base, tier, catID, product

Персистентность: маппинги растут между месяцами (union по listing/PRODUCTID),
новые листинги из выгрузок, которых здесь нет, → на ревью в приложении.
"""
from __future__ import annotations
import os
import pandas as pd

import mp_kinguin, mp_driffle, mp_eneba, mp_g2a, mp_ggsel, mp_plati


def extract_catalog(etalon_paths: list, hdr_map: dict) -> pd.DataFrame:
    """Объединяет листы «Каталог» всех эталонов в один мастер (dedupe по ID)."""
    frames = []
    for p in etalon_paths:
        try:
            c = pd.read_excel(p, sheet_name="Каталог", header=0, engine="calamine")
            c = c[pd.to_numeric(c["ID"], errors="coerce").notna()].copy()
            frames.append(c)
        except Exception:
            pass
    cat = pd.concat(frames, ignore_index=True)
    cat["ID"] = cat["ID"].astype(int)
    # приоритет непустого «Новое назавание»/«Регион»
    cat = cat.sort_values(["ID"]).drop_duplicates("ID", keep="first")
    return cat.reset_index(drop=True)


def extract_mappings(etalons: dict, kinguin_raw: str, out_dir: str = "mappings"):
    os.makedirs(out_dir, exist_ok=True)
    # мастер-каталог
    cat = extract_catalog(list(etalons.values()), {})
    cat.to_csv(f"{out_dir}/catalog_master.csv", index=False)
    print(f"catalog_master.csv: {len(cat)} ID")

    # per-channel
    mp_kinguin.build_kinguin_mapping_from_etalon(etalons["Kinguin"], kinguin_raw) \
        .to_csv(f"{out_dir}/kinguin_mapping.csv", index=False)
    mp_driffle.build_driffle_mapping_from_etalon(etalons["Driffle"]) \
        .to_csv(f"{out_dir}/driffle_mapping.csv", index=False)
    mp_eneba.build_eneba_mapping_from_etalon(etalons["Eneba"]) \
        .to_csv(f"{out_dir}/eneba_mapping.csv", index=False)
    mp_g2a.build_g2a_mapping_from_etalon(etalons["G2A"]) \
        .to_csv(f"{out_dir}/g2a_mapping.csv", index=False)
    mp_plati.build_plati_mapping_from_etalon(etalons["Plati"]) \
        .to_csv(f"{out_dir}/plati_mapping.csv", index=False)
    # GGSel: dict -> csv
    seed = mp_ggsel.build_ggsel_seed(etalons["Plati"])
    rows = [{"base": b, "tier": t, "catID": cid, "product": pr}
            for (b, t), (cid, pr) in seed.items()]
    pd.DataFrame(rows).to_csv(f"{out_dir}/ggsel_mapping.csv", index=False)
    for f in ("kinguin", "driffle", "eneba", "g2a", "plati", "ggsel"):
        n = len(pd.read_csv(f"{out_dir}/{f}_mapping.csv"))
        print(f"{f}_mapping.csv: {n} строк")


# ── загрузчики персистентных маппингов (в родном формате билдеров) ──
def load_df_mapping(path: str) -> pd.DataFrame:
    return pd.read_csv(path)


def load_ggsel_mapping(path: str) -> dict:
    df = pd.read_csv(path)
    return {(str(r["base"]), str(r["tier"])): (int(r["catID"]), r["product"])
            for _, r in df.iterrows()}


if __name__ == "__main__":
    import sys
    base = sys.argv[1] if len(sys.argv) > 1 else "."
    etalons = {
        "Eneba": f"{base}/prodazhi/2605_Продажи на ENEBA_бс.xlsx",
        "Kinguin": f"{base}/prodazhi/2605_Продажи на Kinguin_бс.xlsx",
        "Driffle": f"{base}/prodazhi/2605_Продажи Driffl_бс.xlsx",
        "G2A": f"{base}/prodazhi/2605_Продажи G2A_бс.xlsx",
        "Plati": f"{base}/prodazhi/2605_Продажи на ПЛАТИ_бс.xlsx",
    }
    extract_mappings(etalons, f"{base}/vygruzki/kinguin_202605.xlsx",
                     out_dir=f"{base}/mappings")


# ── персист подтверждённых маппингов из ревью (замыкание цикла обучения) ──
def merge_listing_mapping(current: pd.DataFrame, confirmed: pd.DataFrame) -> pd.DataFrame:
    """Слить подтверждённые [listing, catID, product] в текущий маппинг канала.

    Дедуп по listing (новые/исправленные перекрывают старые). Для каналов,
    ключённых по listing (Eneba/Driffle/G2A/Plati). Возврат — обновлённый DF
    в формате mappings/<channel>_mapping.csv.
    """
    conf = confirmed.dropna(subset=["listing", "catID"]).copy()
    conf["catID"] = pd.to_numeric(conf["catID"], errors="coerce").astype("Int64")
    conf = conf.dropna(subset=["catID"])
    if "product" not in conf.columns:
        conf["product"] = conf["listing"]
    cols = ["listing", "catID", "product"]
    conf = conf[cols]
    both = pd.concat([current[cols] if len(current) else current, conf], ignore_index=True)
    return both.drop_duplicates("listing", keep="last").reset_index(drop=True)


def merge_kinguin_mapping(current: pd.DataFrame, confirmed: pd.DataFrame,
                          kinguin_raw_path: str) -> pd.DataFrame:
    """Слить подтверждённые Kinguin-листинги (ключ PRODUCTID).

    confirmed: [listing (=NAME), catID, product]. PRODUCTID восстанавливается
    из сырья Kinguin по NAME. Дедуп по PRODUCTID. Формат:
    [PRODUCTID, listing, catID, product].
    """
    import mp_kinguin
    raw = mp_kinguin.load_kinguin_raw(kinguin_raw_path) \
        if hasattr(mp_kinguin, "load_kinguin_raw") else \
        pd.read_excel(kinguin_raw_path, sheet_name="reservations",
                      engine="calamine")
    n2p = (raw[["NAME", "PRODUCTID"]].dropna().drop_duplicates("NAME")
              .set_index("NAME")["PRODUCTID"])
    conf = confirmed.dropna(subset=["listing", "catID"]).copy()
    conf["PRODUCTID"] = conf["listing"].map(n2p)
    conf["catID"] = pd.to_numeric(conf["catID"], errors="coerce").astype("Int64")
    conf = conf.dropna(subset=["PRODUCTID", "catID"])
    if "product" not in conf.columns:
        conf["product"] = conf["listing"]
    cols = ["PRODUCTID", "listing", "catID", "product"]
    both = pd.concat([current[cols] if len(current) else current, conf[cols]],
                     ignore_index=True)
    return both.drop_duplicates("PRODUCTID", keep="last").reset_index(drop=True)


def merge_ggsel_mapping(current: pd.DataFrame, confirmed: pd.DataFrame) -> pd.DataFrame:
    """Слить подтверждённые GGSel-листинги (ключ base+tier).

    confirmed: [listing (=Goods), catID, product]. base/tier вычисляются из
    Goods. Дедуп по (base, tier). Формат: [base, tier, catID, product].
    """
    import mp_ggsel
    conf = confirmed.dropna(subset=["listing", "catID"]).copy()
    conf["base"] = conf["listing"].map(mp_ggsel._base)
    conf["tier"] = conf["listing"].map(mp_ggsel._tier)
    conf["catID"] = pd.to_numeric(conf["catID"], errors="coerce").astype("Int64")
    conf = conf.dropna(subset=["catID"])
    if "product" not in conf.columns:
        conf["product"] = conf["listing"]
    cols = ["base", "tier", "catID", "product"]
    both = pd.concat([current[cols] if len(current) else current, conf[cols]],
                     ignore_index=True)
    return both.drop_duplicates(["base", "tier"], keep="last").reset_index(drop=True)
