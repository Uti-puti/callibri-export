"""
export.py — CLI-обёртка для экспорта обращений из Callibri.
Вся бизнес-логика — в core.py.
"""

import os
import sys
import logging
import argparse

sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
load_dotenv()

import core


def main():
    parser = argparse.ArgumentParser(description="Экспорт обращений из Callibri")
    parser.add_argument("--date1", help="Начальная дата (dd.mm.yyyy)")
    parser.add_argument("--date2", help="Конечная дата (dd.mm.yyyy)")
    parser.add_argument("--days", type=int, help="Последние N дней (вместо --date1/--date2)")
    args = parser.parse_args()

    # Настройка логирования
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(
        logging.Formatter("%(asctime)s  %(levelname)s  %(message)s", datefmt="%H:%M:%S")
    )
    logging.basicConfig(level=logging.INFO, handlers=[handler])

    email = os.getenv("CALLIBRI_EMAIL")
    token = os.getenv("CALLIBRI_TOKEN")

    try:
        result = core.run_export(
            email=email,
            token=token,
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
            print(f"  Пропущено чанков    : {result['failed_chunks']} (после {core.MAX_RETRIES} попыток)")
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
