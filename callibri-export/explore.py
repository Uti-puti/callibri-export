"""
explore.py — разведка API Callibri.
Выводит список проектов и пример статистики в консоль.
"""

import os
import sys
import json
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv

# Принудительно выводим UTF-8 в консоль Windows
sys.stdout.reconfigure(encoding="utf-8")

# Загружаем переменные из .env
load_dotenv()

BASE_URL = "https://api.callibri.ru"
EMAIL = os.getenv("CALLIBRI_EMAIL")
TOKEN = os.getenv("CALLIBRI_TOKEN")


def check_credentials():
    """Проверяем что учётные данные загружены из .env"""
    if not EMAIL or not TOKEN or "example" in (EMAIL or ""):
        print("ОШИБКА: заполни CALLIBRI_EMAIL и CALLIBRI_TOKEN в файле .env")
        sys.exit(1)
    print(f"Авторизация: {EMAIL}")
    print()


def auth_params():
    """Общие параметры авторизации для всех запросов"""
    return {"user_email": EMAIL, "user_token": TOKEN}


def get_sites():
    """Шаг 1.1 — получить и вывести список всех проектов"""
    print("=" * 60)
    print("ШАГ 1.1 — Список проектов")
    print("=" * 60)

    resp = requests.get(f"{BASE_URL}/get_sites", params=auth_params())
    resp.raise_for_status()
    data = resp.json()

    # Выводим сырой ответ для отладки (первые 500 символов)
    raw = json.dumps(data, ensure_ascii=False, indent=2)
    print(f"\nСырой ответ (первые 500 символов):\n{raw[:500]}\n")

    sites = data.get("sites", data if isinstance(data, list) else [])
    if not sites:
        print("Проектов не найдено. Проверь токен и email.")
        print(f"Полный ответ: {raw}")
        sys.exit(1)

    print(f"Найдено проектов: {len(sites)}\n")
    print(f"{'site_id':<12} | {'sitename':<30} | {'domains'}")
    print("-" * 80)
    for site in sites:
        site_id = site.get("site_id", "?")
        name = site.get("sitename", "?")
        domains = site.get("domains", [])
        domains_str = ", ".join(domains) if isinstance(domains, list) else str(domains)
        print(f"{site_id:<12} | {name:<30} | {domains_str}")

    print()
    return sites


def get_statistics(site_id, site_name):
    """Шаг 1.2 — статистика по первому проекту за последние 7 дней"""
    print("=" * 60)
    print(f"ШАГ 1.2 — Статистика для: {site_name} (id={site_id})")
    print("=" * 60)

    # Даты: последние 7 дней
    date2 = datetime.now()
    date1 = date2 - timedelta(days=6)  # API считает оба конца включительно, 6 дней = 7-дневный период
    date1_str = date1.strftime("%d.%m.%Y")
    date2_str = date2.strftime("%d.%m.%Y")
    print(f"Период: {date1_str} — {date2_str}\n")

    params = {
        **auth_params(),
        "site_id": site_id,
        "date1": date1_str,
        "date2": date2_str,
    }

    resp = requests.get(f"{BASE_URL}/site_get_statistics", params=params)

    # Показываем тело ответа при ошибке, а не просто HTTP-статус
    if not resp.ok:
        print(f"ОШИБКА {resp.status_code}: {resp.text[:500]}")
        sys.exit(1)

    data = resp.json()

    # Сырой ответ (ключи верхнего уровня)
    print(f"Ключи ответа: {list(data.keys()) if isinstance(data, dict) else type(data)}\n")

    channels = data.get("channels_statistics", [])
    print(f"Количество каналов: {len(channels)}\n")

    # Типы обращений, которые нас интересуют
    appeal_types = ["calls", "feedbacks", "chats", "emails"]

    for ch in channels:
        ch_name = ch.get("name_channel", "—")
        print(f"--- Канал: {ch_name} ---")

        for atype in appeal_types:
            items = ch.get(atype, [])
            print(f"  {atype}: {len(items)} записей")

            # Выводим все поля первой записи
            if items:
                first = items[0]
                print(f"  Поля первой записи ({atype}):")
                for key, val in first.items():
                    # Обрезаем длинные значения
                    val_str = str(val)
                    if len(val_str) > 100:
                        val_str = val_str[:100] + "..."
                    print(f"    {key}: {val_str}")
        print()


def main():
    check_credentials()

    # Шаг 1.1
    sites = get_sites()

    # Шаг 1.2 — берём первый проект
    first = sites[0]
    get_statistics(first.get("site_id"), first.get("sitename", "?"))

    print("Разведка завершена.")


if __name__ == "__main__":
    main()
