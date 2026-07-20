"""Персистентность самообучаемых маппингов через GitHub (репозиторий game-sales).

Streamlit Cloud не пишет в репозиторий из кода напрямую, поэтому подтверждённые
маппинги коммитятся через GitHub Contents API. Чтение НЕ нужно: файлы mappings/
уже лежат на диске деплоя (репозиторий склонирован) — их читает _map_file в app.
После коммита Streamlit Cloud авто-передеплоит, и обновлённый маппинг оказывается
на диске в следующем прогоне.

Секреты (st.secrets):
    [github]
    token  = "github_pat_..."      # fine-grained PAT, Contents: Read and write
    repo   = "username/game-sales" # owner/repo
    branch = "main"                # необязательно (по умолчанию main)
    dir    = "mappings"            # необязательно (по умолчанию mappings)

Все функции мягкие: без requests/токена/репо → False, приложение продолжает
работать локально (маппинги отдаются через ручной zip, как раньше).
"""
from __future__ import annotations

import base64
import json
import pandas as pd

API = "https://api.github.com"

try:
    import requests
    _OK = True
except Exception:
    _OK = False


def cfg_from_secrets(secrets) -> dict | None:
    """Достаёт конфиг из st.secrets['github']; None если секции нет."""
    try:
        g = secrets["github"]
        return {"token": g["token"], "repo": g["repo"],
                "branch": g.get("branch", "main"),
                "dir": g.get("dir", "mappings")}
    except Exception:
        return None


def available(cfg: dict | None) -> bool:
    return bool(_OK and cfg and cfg.get("token") and cfg.get("repo"))


def repo_label(cfg: dict | None) -> str:
    return f"{cfg['repo']}@{cfg['branch']}" if available(cfg) else ""


def _headers(cfg: dict) -> dict:
    return {"Authorization": f"Bearer {cfg['token']}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28"}


def _get_sha(cfg: dict, path: str):
    r = requests.get(f"{API}/repos/{cfg['repo']}/contents/{path}",
                     params={"ref": cfg["branch"]}, headers=_headers(cfg), timeout=20)
    return r.json().get("sha") if r.status_code == 200 else None


def save_mapping(cfg: dict | None, name: str, df: pd.DataFrame) -> bool:
    """Создать/обновить mappings/{name}_mapping.csv в репозитории. True при успехе."""
    if not available(cfg):
        return False
    try:
        path = f"{cfg['dir']}/{name.lower()}_mapping.csv"
        content = base64.b64encode(df.to_csv(index=False).encode("utf-8")).decode()
        body = {"message": f"SalesFlow: обновление {name.lower()}_mapping.csv",
                "content": content, "branch": cfg["branch"]}
        sha = _get_sha(cfg, path)
        if sha:
            body["sha"] = sha           # обновление существующего файла
        r = requests.put(f"{API}/repos/{cfg['repo']}/contents/{path}",
                         headers=_headers(cfg), data=json.dumps(body), timeout=30)
        return r.status_code in (200, 201)
    except Exception:
        return False


def save_many(cfg: dict | None, updated: dict) -> list:
    """Закоммитить несколько маппингов. Возврат — имена успешно сохранённых."""
    return [n for n, df in updated.items() if save_mapping(cfg, n, df)]


def check(cfg: dict | None) -> tuple:
    """(ok, сообщение) — быстрая проверка доступа к репозиторию."""
    if not available(cfg):
        return False, "GitHub не настроен"
    try:
        r = requests.get(f"{API}/repos/{cfg['repo']}", headers=_headers(cfg), timeout=20)
        if r.status_code == 200:
            perms = r.json().get("permissions", {})
            if perms.get("push"):
                return True, f"{cfg['repo']}@{cfg['branch']} — запись доступна"
            return False, "нет прав на запись (нужен Contents: Read and write)"
        return False, f"репозиторий недоступен (HTTP {r.status_code})"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
