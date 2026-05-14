"""
accounts.py — именованные аккаунты провайдеров (Calltouch и др.).

Аккаунты позволяют выгружать разные сайты под разными API-ключами
(когда у одного оператора несколько Calltouch-аккаунтов с разными правами).

Структура accounts.json:
    {
      "calltouch": [
        {"name": "Kazna", "client_api_id": "..."},
        {"name": "d-okna", "client_api_id": "..."}
      ]
    }

Если у проекта в projects.json задан "account": "<name>" — run_export
возьмёт креды этого аккаунта. Иначе — креды из .env (дефолт).
"""

import json
import os


def _default_path():
    import core
    return os.path.join(core.get_app_dir(), "accounts.json")


def load_accounts(path=None):
    """Загрузить accounts.json. Если файла нет — вернуть пустой dict."""
    if path is None:
        path = _default_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    if not isinstance(data, dict):
        return {}
    return data


def save_accounts(data, path=None):
    """Сохранить accounts.json."""
    if path is None:
        path = _default_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_provider_accounts(provider_name, accounts_data=None):
    """Список аккаунтов для провайдера: [{name, ...}, ...]."""
    if accounts_data is None:
        accounts_data = load_accounts()
    items = accounts_data.get(provider_name) or []
    return [a for a in items if isinstance(a, dict) and a.get("name")]


def get_account(provider_name, account_name, accounts_data=None):
    """Найти аккаунт по имени. Вернуть dict или None."""
    if not account_name:
        return None
    for a in get_provider_accounts(provider_name, accounts_data):
        if a.get("name") == account_name:
            return a
    return None


def account_credentials(account):
    """Креды без служебного поля name."""
    if not account:
        return {}
    return {k: v for k, v in account.items() if k != "name"}
