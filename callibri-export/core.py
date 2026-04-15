"""
core.py — бизнес-логика экспорта обращений из Callibri.
Чистый модуль без CLI, без sys.exit. Все ошибки — через исключения.
Используется и CLI-обёрткой (export.py), и GUI (app.py).
"""

import csv
import os
import re
import sys
import json
import time
import logging
from datetime import datetime, timedelta

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

log = logging.getLogger(__name__)

BASE_URL = "https://api.callibri.ru"

# Типы обращений
APPEAL_TYPES = ["calls", "feedbacks", "chats", "emails"]

# Колонки по умолчанию
DEFAULT_COLUMNS = [
    "date", "name_channel", "comment", "status", "type",
    "conversations_number", "utm_campaign",
]

MAX_RETRIES = 3  # максимум попыток на один чанк


# ── Утилиты ──────────────────────────────────────────────────────────────────

def get_app_dir():
    """Директория приложения (рядом с .exe или рядом с .py)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def parse_date(value, name="date"):
    """Парсим дату из строки dd.mm.yyyy. Бросает ValueError при ошибке."""
    try:
        return datetime.strptime(value, "%d.%m.%Y")
    except ValueError:
        raise ValueError(
            f"Неверный формат даты {name}='{value}'. Ожидается dd.mm.yyyy (например, 01.03.2026)"
        )


def resolve_period(date1_str=None, date2_str=None, days=None):
    """Определяем период выгрузки. Возвращает (date1, date2)."""
    if date1_str and date2_str:
        date1 = parse_date(date1_str, "date1")
        date2 = parse_date(date2_str, "date2")
        if date1 > date2:
            raise ValueError(
                f"Начальная дата ({date1_str}) позже конечной ({date2_str})"
            )
    elif days:
        if days < 1:
            raise ValueError(f"days должен быть >= 1, получено: {days}")
        date2 = datetime.now()
        date1 = date2 - timedelta(days=days - 1)
    else:
        date2 = datetime.now()
        date1 = date2 - timedelta(days=6)
    return date1, date2


def split_period(date1, date2):
    """Разбиваем период на чанки по 7 дней включительно."""
    chunks = []
    current = date1
    while current <= date2:
        chunk_end = min(current + timedelta(days=6), date2)
        chunks.append((current, chunk_end))
        current = chunk_end + timedelta(days=1)
    return chunks


def format_date(iso_date_str):
    """Конвертируем ISO дату в читаемый формат dd.mm.yyyy HH:MM"""
    if not iso_date_str:
        return ""
    try:
        dt = datetime.strptime(iso_date_str[:19], "%Y-%m-%dT%H:%M:%S")
        return dt.strftime("%d.%m.%Y %H:%M")
    except (ValueError, TypeError):
        return str(iso_date_str)


def extract_appeal_id(appeal):
    """appeal_id из записи, или clbvid как запасной вариант (для дедупликации)."""
    aid = appeal.get("appeal_id")
    if aid:
        return str(aid)
    clbvid = appeal.get("clbvid")
    return str(clbvid) if clbvid else ""


def sanitize_filename(name):
    """Убираем из имени канала символы, недопустимые в именах файлов."""
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


# ── Конфигурация ─────────────────────────────────────────────────────────────

def check_credentials(email, token):
    """Проверяем учётные данные. Возвращает (ok, message)."""
    if not email or not token:
        return False, "Заполни email и token"
    if "example" in email:
        return False, "Замени example-email на реальный"
    return True, "OK"


def load_projects(path=None):
    """Загружаем список проектов из projects.json.
    Бросает FileNotFoundError / ValueError при проблемах."""
    if path is None:
        path = os.path.join(get_app_dir(), "projects.json")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Файл projects.json не найден: {path}\n"
            'Создай его рядом с приложением. Пример: [{"site_id": 4112, "folder": "energocenter"}]'
        )
    with open(path, encoding="utf-8") as f:
        projects = json.load(f)
    if not projects:
        raise ValueError("projects.json пустой — добавь хотя бы один проект.")
    return projects


def save_projects(projects, path=None):
    """Сохраняем список проектов в projects.json."""
    if path is None:
        path = os.path.join(get_app_dir(), "projects.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)


# Все известные поля: имя → описание
FIELD_DESCRIPTIONS = {
    # Специальные (вычисляемые)
    "date": "Дата обращения (dd.mm.yyyy HH:MM)",
    "name_channel": "Название канала",
    "type": "Тип обращения (calls/feedbacks/chats/emails)",
    # API-поля
    "appeal_id": "ID обращения (уникальный)",
    "phone": "Телефон клиента",
    "email": "E-mail клиента",
    "name": "Имя клиента",
    "comment": "Комментарий к обращению",
    "content": "Полный текст обращения (заявки/email)",
    "status": "Статус обращения (Лид, Нет ответа...)",
    "source": "Источник трафика (Google, Yandex...)",
    "traffic_type": "Тип трафика",
    "region": "Регион клиента",
    "device": "Устройство (desktop/mobile/tablet)",
    "conversations_number": "Номер обращения клиента (1 = первое)",
    "is_lid": "Является ли лидом (True/False)",
    "name_type": "Тип обращения (Звонок, Заявка, E-mail)",
    "landing_page": "Страница входа",
    "lid_landing": "Целевая страница лида",
    "site_referrer": "Реферер",
    "link_download": "Ссылка на запись звонка",
    "duration": "Длительность звонка, сек (только звонки)",
    "billsec": "Длительность разговора, сек (только звонки)",
    "responsible_manager": "Ответственный менеджер",
    "responsible_manager_email": "Email менеджера",
    "call_status": "Статус звонка",
    "accurately": "Точное определение источника (True/False)",
    "form_name": "Название формы (только заявки)",
    "utm_source": "UTM source",
    "utm_medium": "UTM medium",
    "utm_campaign": "UTM campaign",
    "utm_content": "UTM content",
    "utm_term": "UTM term",
    "query": "Поисковый запрос",
    "channel_id": "ID канала",
    "crm_client_id": "ID клиента в CRM",
    "ym_uid": "Яндекс.Метрика UID",
    "metrika_client_id": "Яндекс.Метрика Client ID",
    "ua_client_id": "Google Analytics Client ID",
    "clbvid": "Внутренний ID Callibri",
}

# Список всех полей (порядок)
ALL_FIELDS = list(FIELD_DESCRIPTIONS.keys())


# ── API ──────────────────────────────────────────────────────────────────────

def auth_params(email, token):
    """Общие параметры авторизации для всех запросов."""
    return {"user_email": email, "user_token": token}


def get_sites(email, token):
    """Получить список всех проектов из API."""
    resp = requests.get(
        f"{BASE_URL}/get_sites", params=auth_params(email, token), timeout=30
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("sites", [])


def get_statistics(site_id, date1_str, date2_str, email, token):
    """Получить статистику по проекту за указанный период (макс 7 дней)."""
    params = {
        **auth_params(email, token),
        "site_id": site_id,
        "date1": date1_str,
        "date2": date2_str,
    }
    resp = requests.get(
        f"{BASE_URL}/site_get_statistics", params=params, timeout=60
    )
    resp.raise_for_status()
    return resp.json()


def test_connection(email, token):
    """Проверка подключения к API. Возвращает (success, site_count, message)."""
    try:
        ok, msg = check_credentials(email, token)
        if not ok:
            return False, 0, msg
        sites = get_sites(email, token)
        return True, len(sites), f"Подключено. Проектов: {len(sites)}"
    except Exception as e:
        return False, 0, f"Ошибка подключения: {e}"


def get_channels_and_statuses(site_id, email, token):
    """Запрашиваем статистику за последние 7 дней и извлекаем имена каналов и статусы.
    7 дней — максимальный период одного запроса API, даёт надёжный результат
    даже если в последние 1-2 дня обращений не было.
    Возвращает (channel_names: list[str], statuses: list[str])."""
    today = datetime.now()
    week_ago = today - timedelta(days=6)
    d1 = week_ago.strftime("%d.%m.%Y")
    d2 = today.strftime("%d.%m.%Y")
    data = get_statistics(site_id, d1, d2, email, token)

    channels = data.get("channels_statistics", [])
    channel_names = []
    statuses = set()
    for ch in channels:
        name = ch.get("name_channel", "")
        if name:
            channel_names.append(name)
        for atype in APPEAL_TYPES:
            for appeal in ch.get(atype, []):
                s = appeal.get("status", "")
                if s:
                    statuses.add(s)
    return channel_names, sorted(statuses)


# ── Парсинг данных ───────────────────────────────────────────────────────────

def build_row(appeal, ch_name, atype, columns):
    """Собираем строку из записи обращения по списку запрошенных колонок."""
    row = {}
    for col in columns:
        if col == "date":
            row[col] = format_date(appeal.get("date"))
        elif col == "name_channel":
            row[col] = ch_name
        elif col == "type":
            row[col] = atype
        else:
            row[col] = appeal.get(col, "") or ""
    return row


def parse_chunk_data(data, channel_filter, seen_ids, columns=None,
                     type_filter=None, status_filter=None):
    """
    Парсим ответ одного чанка в rows_by_channel.
    Возвращает dict {channel_name: [rows]} и кол-во записей в чанке.
    """
    if columns is None:
        columns = DEFAULT_COLUMNS

    appeal_types = type_filter or APPEAL_TYPES
    channels = data.get("channels_statistics", [])
    rows_by_channel = {}
    chunk_count = 0

    for channel in channels:
        ch_name = channel.get("name_channel", "")
        if channel_filter and ch_name not in channel_filter:
            continue

        channel_rows = []
        for atype in appeal_types:
            appeals = channel.get(atype, [])
            if not appeals:
                continue
            for appeal in appeals:
                appeal_id = extract_appeal_id(appeal)
                if appeal_id and appeal_id in seen_ids:
                    continue
                if appeal_id:
                    seen_ids.add(appeal_id)
                if status_filter and (appeal.get("status", "") or "") not in status_filter:
                    continue
                channel_rows.append(build_row(appeal, ch_name, atype, columns))

        if channel_rows:
            rows_by_channel[ch_name] = channel_rows
            chunk_count += len(channel_rows)

    return rows_by_channel, chunk_count


# ── Запросы с retry ──────────────────────────────────────────────────────────

def _emit(on_log, message):
    """Отправить сообщение в callback или в logging."""
    if on_log:
        on_log(message)
    else:
        log.info(message)


def fetch_chunk_with_retry(site_id, d1, d2, idx, total_chunks,
                           email, token, on_log=None):
    """Запрос чанка с повторными попытками. Возвращает data или None."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            data = get_statistics(site_id, d1, d2, email, token)
            return data
        except requests.HTTPError as e:
            _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): HTTP ошибка — {e} (попытка {attempt}/{MAX_RETRIES})")
        except requests.RequestException as e:
            _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): сетевая ошибка — {e} (попытка {attempt}/{MAX_RETRIES})")

        if attempt < MAX_RETRIES:
            pause = attempt * 2
            _emit(on_log, f"    Повторная попытка через {pause}с...")
            time.sleep(pause)

    _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): все {MAX_RETRIES} попытки исчерпаны — пропускаем")
    return None


# ── Обработка проекта ────────────────────────────────────────────────────────

def process_site(site, chunks, email, token, channel_filter=None, columns=None,
                 type_filter=None, status_filter=None, on_log=None, on_chunk=None):
    """
    Обработка одного проекта за весь период.
    on_chunk(current_chunk, total_chunks) — callback прогресса по чанкам.
    Возвращает (успех, {channel_name: [rows]}, failed_chunks).
    """
    site_id = site.get("site_id")
    total_chunks = len(chunks)

    merged = {}
    seen_ids = set()
    has_data = False
    failed_chunks = 0

    for idx, (chunk_start, chunk_end) in enumerate(chunks, start=1):
        d1 = chunk_start.strftime("%d.%m.%Y")
        d2 = chunk_end.strftime("%d.%m.%Y")

        data = fetch_chunk_with_retry(site_id, d1, d2, idx, total_chunks,
                                      email, token, on_log)
        if data is None:
            failed_chunks += 1
            if on_chunk:
                on_chunk(idx, total_chunks)
            continue

        chunk_rows, chunk_count = parse_chunk_data(
            data, channel_filter, seen_ids, columns, type_filter, status_filter
        )

        for ch_name, rows in chunk_rows.items():
            merged.setdefault(ch_name, []).extend(rows)

        has_data = True
        _emit(on_log, f"    Чанк {idx}/{total_chunks}: {d1} — {d2} — {chunk_count} записей")

        if on_chunk:
            on_chunk(idx, total_chunks)

        if idx < total_chunks:
            time.sleep(0.5)

    return has_data, merged, failed_chunks


# ── Запись файлов ────────────────────────────────────────────────────────────

def write_csv(filepath, rows, columns=None):
    """Записываем CSV с заголовком (разделитель ;, UTF-8 BOM)."""
    if columns is None:
        columns = DEFAULT_COLUMNS
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore", delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(filepath, rows, columns=None):
    """Записываем XLSX с заголовком и автошириной колонок."""
    if columns is None:
        columns = DEFAULT_COLUMNS

    wb = Workbook()
    ws = wb.active

    header_font = Font(bold=True)
    header_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")

    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 50)

    wb.save(filepath)


# ── Главная функция экспорта ─────────────────────────────────────────────────

def run_export(
    email, token,
    date1_str=None, date2_str=None, days=None,
    projects_path=None,
    output_dir=None,
    enabled_site_ids=None,
    on_log=None,
    on_progress=None,
):
    """
    Главная функция экспорта.

    Параметры:
    - email, token: учётные данные Callibri
    - date1_str, date2_str: период (dd.mm.yyyy), или days для последних N дней
    - projects_path: путь к projects.json (None = рядом с приложением)
    - output_dir: папка для результатов (None = output/ рядом с приложением)
    - enabled_site_ids: set of site_id (если передан — переопределяет enabled из json)
    - on_log(message): callback для логирования
    - on_progress(current_project, total_projects, current_chunk, total_chunks):
      callback для прогресса

    Бросает ValueError при ошибках валидации.
    Возвращает dict с результатами.
    """
    # 1. Проверка учётных данных
    ok, msg = check_credentials(email, token)
    if not ok:
        raise ValueError(msg)

    # 2. Определяем период
    date1, date2 = resolve_period(date1_str, date2_str, days)
    chunks = split_period(date1, date2)
    _emit(on_log, f"Период: {date1.strftime('%d.%m.%Y')} — {date2.strftime('%d.%m.%Y')} ({len(chunks)} чанков)")

    # 3. Загружаем конфигурацию
    projects_config = load_projects(projects_path)
    _emit(on_log, f"Проектов в конфиге: {len(projects_config)}")

    # 4. Загружаем проекты из API
    _emit(on_log, "Загружаем список проектов из API...")
    all_sites = get_sites(email, token)
    sites_lookup = {s["site_id"]: s for s in all_sites}

    if output_dir is None:
        output_dir = os.path.join(get_app_dir(), "output")
    date_suffix = f"callibri_{date1.strftime('%d%m%Y')}-{date2.strftime('%d%m%Y')}"

    # 5. Предпроверка
    active_projects = []
    for p in projects_config:
        sid = p.get("site_id")
        if enabled_site_ids is not None:
            enabled = sid in enabled_site_ids
        else:
            enabled = p.get("enabled", True)
        if enabled:
            active_projects.append(p)

    missing = [p for p in active_projects if p.get("site_id") not in sites_lookup]
    if missing:
        for p in missing:
            _emit(on_log, f"ПРЕДУПРЕЖДЕНИЕ: site_id={p.get('site_id')} (folder={p.get('folder')}) не найден в API")
        _emit(on_log, f"Не найдено проектов: {len(missing)} из {len(active_projects)} активных")

    # 6. Цикл по проектам
    processed = 0
    disabled = 0
    errors = 0
    total_failed_chunks = 0
    report = []
    total_active = len(active_projects)

    for i, proj_conf in enumerate(projects_config):
        site_id = proj_conf.get("site_id")
        folder = proj_conf.get("folder")
        channel_filter = proj_conf.get("channels")
        split_by_channel = proj_conf.get("split_by_channel", False)
        columns = proj_conf.get("fields") or DEFAULT_COLUMNS
        type_filter = proj_conf.get("types")
        status_filter = proj_conf.get("statuses")
        out_format = proj_conf.get("format", "xlsx")

        # Определяем enabled
        if enabled_site_ids is not None:
            enabled = site_id in enabled_site_ids
        else:
            enabled = proj_conf.get("enabled", True)

        if not enabled:
            disabled += 1
            continue

        if site_id not in sites_lookup:
            errors += 1
            continue

        site = sites_lookup[site_id]
        site_name = site.get("sitename", str(site_id))
        _emit(on_log, f"Обрабатываем: {site_name} (id={site_id}) → output/{folder}/")

        if channel_filter:
            _emit(on_log, f"  Фильтр каналов: {channel_filter}")
        if type_filter:
            _emit(on_log, f"  Фильтр типов: {type_filter}")
        if status_filter:
            _emit(on_log, f"  Фильтр статусов: {status_filter}")
        if columns != DEFAULT_COLUMNS:
            _emit(on_log, f"  Поля: {columns}")

        # Callback прогресса по чанкам
        project_idx = processed  # номер среди активных
        def _on_chunk(cur_chunk, tot_chunks, _pi=project_idx):
            if on_progress:
                on_progress(_pi, total_active, cur_chunk, tot_chunks)

        ok_data, rows_by_channel, failed_chunks = process_site(
            site, chunks, email, token, channel_filter, columns,
            type_filter, status_filter, on_log, _on_chunk
        )
        total_failed_chunks += failed_chunks

        if not ok_data:
            errors += 1
            report.append((site_name, 0))
            continue

        project_dir = os.path.join(output_dir, folder)
        os.makedirs(project_dir, exist_ok=True)

        write_fn = write_csv if out_format == "csv" else write_xlsx
        ext = out_format

        if split_by_channel:
            project_total = 0
            for ch_name, rows in rows_by_channel.items():
                rows.sort(key=lambda r: r.get("date", ""))
                safe_name = sanitize_filename(ch_name)
                filepath = os.path.join(project_dir, f"{date_suffix}_{safe_name}.{ext}")
                write_fn(filepath, rows, columns)
                _emit(on_log, f"  [{ch_name}] → {len(rows)} строк → {os.path.basename(filepath)}")
                project_total += len(rows)
        else:
            all_rows = [row for rows in rows_by_channel.values() for row in rows]
            all_rows.sort(key=lambda r: r.get("date", ""))
            filepath = os.path.join(project_dir, f"{date_suffix}.{ext}")
            write_fn(filepath, all_rows, columns)
            _emit(on_log, f"  Сохранено: {len(all_rows)} строк → {os.path.basename(filepath)}")
            project_total = len(all_rows)

        if failed_chunks:
            _emit(on_log, f"  Внимание: {failed_chunks} чанк(ов) не удалось загрузить — данные неполные")

        report.append((site_name, project_total))
        processed += 1

    # Итог
    result = {
        "processed": processed,
        "disabled": disabled,
        "errors": errors,
        "failed_chunks": total_failed_chunks,
        "report": report,
    }

    summary = f"Готово! Обработано: {processed}"
    if report:
        total_rows = sum(c for _, c in report)
        summary += f", строк: {total_rows}"
    _emit(on_log, summary)

    return result
