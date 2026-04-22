"""
explore.py — разведка API провайдера.
Выводит список проектов и пример полей первой записи.

    python explore.py                    # callibri (default)
    python explore.py --provider calltouch
"""

import os
import sys
import json
import argparse
from datetime import datetime, timedelta

from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv()

import providers


def _build_credentials(provider):
    """Собрать creds для провайдера из .env через CREDENTIAL_FIELDS."""
    creds = {}
    for field in getattr(provider, "CREDENTIAL_FIELDS", []):
        env_name = field.get("env")
        if env_name:
            creds[field["key"]] = os.getenv(env_name, "")
    return creds


def dump_sites(provider, creds):
    """Список проектов."""
    print("=" * 60)
    print(f"Проекты [{provider.LABEL}]")
    print("=" * 60)
    sites = provider.list_sites(creds)
    if not sites:
        print("Проектов не найдено.")
        sys.exit(1)

    print(f"Найдено: {len(sites)}\n")
    print(f"{'site_id':<12} | {'name':<30} | {'domains'}")
    print("-" * 80)
    for site in sites:
        sid = site.get("site_id", "?")
        name = site.get("sitename") or site.get("name", "?")
        domains = site.get("domains", [])
        domains_str = ", ".join(domains) if isinstance(domains, list) else str(domains)
        print(f"{str(sid):<12} | {str(name):<30} | {domains_str}")
    print()
    return sites


def dump_sample(provider, creds, site):
    """Получить данные за 7 дней и показать поля первой записи."""
    print("=" * 60)
    sid = site.get("site_id")
    name = site.get("sitename") or site.get("name", "?")
    print(f"Пример данных: {name} (id={sid})")
    print("=" * 60)

    date2 = datetime.now()
    date1 = date2 - timedelta(days=6)
    print(f"Период: {date1.strftime('%d.%m.%Y')} — {date2.strftime('%d.%m.%Y')}\n")

    filters = {"columns": list(provider.ALL_FIELDS)}
    has_data, rows_by_channel, failed = provider.process_site(
        site, date1, date2, creds, filters,
    )

    if not has_data:
        print("Нет данных за период.")
        return

    print(f"Каналов/групп с данными: {len(rows_by_channel)}\n")
    for ch_name, rows in rows_by_channel.items():
        print(f"--- {ch_name}: {len(rows)} записей ---")
        if rows:
            print("  Поля первой записи:")
            for key, val in rows[0].items():
                val_str = str(val)
                if len(val_str) > 100:
                    val_str = val_str[:100] + "..."
                print(f"    {key}: {val_str}")
        print()


def main():
    parser = argparse.ArgumentParser(description="Разведка API провайдера")
    parser.add_argument(
        "--provider", default="callibri",
        choices=providers.provider_names(),
        help="Имя провайдера (callibri / calltouch)",
    )
    parser.add_argument(
        "--no-sample", action="store_true",
        help="Только список проектов, без выборки данных",
    )
    args = parser.parse_args()

    provider = providers.get_provider(args.provider)
    creds = _build_credentials(provider)

    ok, msg = provider.check_credentials(creds)
    if not ok:
        print(f"ОШИБКА [{provider.LABEL}]: {msg}")
        sys.exit(1)

    print(f"Авторизация: {provider.LABEL}\n")

    sites = dump_sites(provider, creds)
    if not args.no_sample and sites:
        dump_sample(provider, creds, sites[0])

    print("Разведка завершена.")


if __name__ == "__main__":
    main()
