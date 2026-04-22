"""
gsheets.py — автономный модуль для экспорта данных в Google Sheets.

Не импортирует core.py. Получает готовые данные (list[dict] + columns) и отправляет
в указанную таблицу. Можно использовать отдельно от приложения.

Подход: Service Account (файл credentials.json) через gspread + google-auth.
"""

import json
import re
import time
import logging

log = logging.getLogger(__name__)

# Ленивый импорт — gspread/google-auth нужны только при реальном использовании.
# Это позволяет приложению работать без этих пакетов, если GSheets не настроен.
_gspread = None
_ServiceAccountCredentials = None


def _ensure_imports():
    """Ленивый импорт gspread и google-auth. Бросает ImportError при отсутствии."""
    global _gspread, _ServiceAccountCredentials
    if _gspread is None:
        import gspread
        from google.oauth2.service_account import Credentials
        _gspread = gspread
        _ServiceAccountCredentials = Credentials


# ── Авторизация ─────────────────────────────────────────────────────────────


def get_service_account_email(credentials_path):
    """Извлекает email сервисного аккаунта из JSON-ключа (без авторизации)."""
    with open(credentials_path, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("client_email", "")


def authorize(credentials_path):
    """
    Авторизация через Service Account.
    Возвращает gspread.Client.
    Бросает FileNotFoundError, ValueError, google.auth exceptions.
    """
    _ensure_imports()
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = _ServiceAccountCredentials.from_service_account_file(
        credentials_path, scopes=scopes
    )
    client = _gspread.authorize(creds)
    return client


def test_gsheet_connection(credentials_path):
    """
    Проверка подключения к Google Sheets.
    Возвращает (success, sa_email, message).
    """
    try:
        sa_email = get_service_account_email(credentials_path)
        client = authorize(credentials_path)
        # Пробуем выполнить запрос — список таблиц
        client.list_spreadsheet_files(title=None)
        return True, sa_email, f"Подключено. Service Account: {sa_email}"
    except FileNotFoundError:
        return False, "", f"Файл не найден: {credentials_path}"
    except ImportError:
        return False, "", "Не установлены пакеты gspread / google-auth"
    except Exception as e:
        return False, "", f"Ошибка подключения: {e}"


# ── Работа с таблицами и листами ────────────────────────────────────────────


def parse_spreadsheet_id(url_or_id):
    """
    Извлекает spreadsheet_id из URL или возвращает как есть если это ID.
    URL формат: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit...
    """
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url_or_id)
    if match:
        return match.group(1)
    # Предполагаем что это уже ID (длинная строка без пробелов)
    stripped = url_or_id.strip()
    if stripped and " " not in stripped:
        return stripped
    raise ValueError(f"Не удалось извлечь ID таблицы из: {url_or_id}")


def list_spreadsheets(client):
    """
    Список таблиц, доступных сервисному аккаунту.
    Возвращает list[dict] с ключами: id, title.
    """
    files = client.list_spreadsheet_files()
    return [{"id": f["id"], "title": f["name"]} for f in files]


def get_spreadsheet_info(client, spreadsheet_id):
    """
    Информация о таблице: название и список листов.
    Возвращает (title, [sheet_names]).
    """
    spreadsheet = client.open_by_key(spreadsheet_id)
    title = spreadsheet.title
    sheet_names = [ws.title for ws in spreadsheet.worksheets()]
    return title, sheet_names


def create_sheet(client, spreadsheet_id, sheet_name):
    """Создать новый лист (вкладку) в таблице. Возвращает имя листа."""
    spreadsheet = client.open_by_key(spreadsheet_id)
    spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=26)
    return sheet_name


# ── Экспорт данных ──────────────────────────────────────────────────────────


def export_to_sheet(client, spreadsheet_id, sheet_name, rows, columns,
                    mode="append", on_log=None):
    """
    Основная функция записи данных в Google Sheets.

    Параметры:
    - client: авторизованный gspread.Client
    - spreadsheet_id: ID таблицы
    - sheet_name: имя листа
    - rows: list[dict] — данные (ключи = имена полей)
    - columns: list[str] — порядок колонок
    - mode: "append" (дополнить) или "replace" (заменить)
    - on_log: callback для сообщений

    Возвращает dict: {"rows_written": N, "start_row": N, "url": str}
    """
    def _log(msg):
        if on_log:
            on_log(f"Google Sheets: {msg}")
        else:
            log.info(f"Google Sheets: {msg}")

    if not rows:
        _log("нет данных для записи")
        return {"rows_written": 0, "start_row": 0, "url": ""}

    spreadsheet = client.open_by_key(spreadsheet_id)
    _log(f'подключение к "{spreadsheet.title}"...')

    # Открываем лист
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except _gspread.exceptions.WorksheetNotFound:
        raise ValueError(f'Лист "{sheet_name}" не найден в таблице "{spreadsheet.title}"')

    # Готовим данные как список списков
    header = list(columns)
    data_rows = []
    for row in rows:
        data_rows.append([str(row.get(col, "") or "") for col in columns])

    if mode == "replace":
        _log(f'лист "{sheet_name}" — режим "заменить"')
        worksheet.clear()
        time.sleep(0.3)

        # Записываем заголовок + данные одним batch
        all_values = [header] + data_rows
        _batch_update(worksheet, all_values, start_row=1, on_log=_log)

        start_row = 2
        _log(f"очищен лист, записан заголовок + {len(data_rows)} строк")

    elif mode == "append":
        _log(f'лист "{sheet_name}" — режим "дополнить"')

        # Читаем текущие данные чтобы найти последнюю строку и проверить заголовок
        existing = worksheet.get_all_values()

        if not existing:
            # Лист пустой — пишем заголовок + данные
            all_values = [header] + data_rows
            _batch_update(worksheet, all_values, start_row=1, on_log=_log)
            start_row = 2
            _log(f"лист был пуст, записан заголовок + {len(data_rows)} строк")
        else:
            # Проверяем заголовок
            existing_header = existing[0] if existing else []
            if existing_header != header:
                _log("ПРЕДУПРЕЖДЕНИЕ — заголовок в таблице отличается от текущих полей")

            # Находим последнюю заполненную строку
            last_row = len(existing)
            start_row = last_row + 1
            _log(f"найдена последняя строка: {last_row}")

            # Дозаписываем данные
            _batch_update(worksheet, data_rows, start_row=start_row, on_log=_log)
            _log(f"записано {len(data_rows)} строк (строки {start_row}–{start_row + len(data_rows) - 1})")

    else:
        raise ValueError(f"Неизвестный режим записи: {mode}. Ожидается 'append' или 'replace'")

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    _log(f"готово — {url}")

    return {
        "rows_written": len(data_rows),
        "start_row": start_row,
        "url": url,
    }


# ── Вспомогательные ─────────────────────────────────────────────────────────


# Google Sheets API квота: 60 запросов/мин. Пишем пакетами по 500 строк.
BATCH_SIZE = 500


def _batch_update(worksheet, values, start_row=1, on_log=None):
    """
    Записывает данные пакетами, чтобы не превысить квоту API.
    values — list[list[str]], start_row — номер строки (1-based).
    """
    if not values:
        return

    total = len(values)
    for offset in range(0, total, BATCH_SIZE):
        batch = values[offset:offset + BATCH_SIZE]
        row_start = start_row + offset
        row_end = row_start + len(batch) - 1
        col_end = len(batch[0]) if batch else 1

        # A1-нотация для диапазона
        end_col_letter = _col_letter(col_end)
        cell_range = f"A{row_start}:{end_col_letter}{row_end}"

        worksheet.update(cell_range, batch, value_input_option="RAW")

        # Пауза между пакетами для квоты API
        if offset + BATCH_SIZE < total:
            time.sleep(1.0)


def _col_letter(col_num):
    """Номер колонки (1-based) → буква A-Z, AA-AZ и т.д."""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result
