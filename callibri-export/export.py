"""
export.py — CLI-обёртка для экспорта обращений.
Вся бизнес-логика — в core.py и providers/.
"""

import os
import sys
import logging
import argparse

sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
load_dotenv()

import core
import providers


def _build_credentials():
    """Собрать dict учётных данных из переменных окружения.

    Для каждого провайдера подтягиваются переменные, описанные в его
    CREDENTIAL_FIELDS (ключ "env" → имя переменной окружения).
    """
    creds = {}
    for prov in providers.all_providers():
        prov_creds = {}
        for field in getattr(prov, "CREDENTIAL_FIELDS", []):
            env_name = field.get("env")
            if env_name:
                prov_creds[field["key"]] = os.getenv(env_name, "")
        creds[prov.NAME] = prov_creds
    return creds


def main():
    parser = argparse.ArgumentParser(description="Экспорт обращений (Callibri / Calltouch)")
    parser.add_argument("--date1", help="Начальная дата (dd.mm.yyyy)")
    parser.add_argument("--date2", help="Конечная дата (dd.mm.yyyy)")
    parser.add_argument("--days", type=int, help="Последние N дней (вместо --date1/--date2)")
    args = parser.parse_args()

    # Логирование
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(
        logging.Formatter("%(asctime)s  %(levelname)s  %(message)s", datefmt="%H:%M:%S")
    )
    logging.basicConfig(level=logging.INFO, handlers=[handler])

    credentials = _build_credentials()

    try:
        result = core.run_export(
            credentials=credentials,
            date1_str=args.date1,
            date2_str=args.date2,
            days=args.days,
        )

        print()
        print("=" * 50)
        print("Готово!")
        print(f"  Обработано проектов : {result['processed']}")
        if result["disabled"]:
            print(f"  Выключено           : {result['disabled']}")
        if result["errors"]:
            print(f"  Ошибки              : {result['errors']}")
        if result["failed_chunks"]:
            print(f"  Пропущено чанков    : {result['failed_chunks']}")
        if result["report"]:
            print()
            for name, count in result["report"]:
                print(f"  {name:<30} {count} строк")
        print("=" * 50)

    except (ValueError, FileNotFoundError, ConnectionError) as e:
        print(f"ОШИБКА: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
