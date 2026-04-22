"""
providers/callibri.py — провайдер Callibri API.

Инкапсулирует всю специфику Callibri: авторизация query-параметрами,
7-дневные чанки, структура channels_statistics, поля API.

Credentials: {"email": str, "token": str}
"""

import re
import time
import logging
from datetime import datetime, timedelta

import requests

log = logging.getLogger(__name__)

# Маскируем секреты в query-параметрах перед попаданием URL в логи/ошибки.
_SECRET_QUERY_KEYS = ("user_token", "user_email", "token", "apiKey", "api_key")
_REDACT_RE = re.compile(
    r"(?i)(" + "|".join(_SECRET_QUERY_KEYS) + r")=[^&\s]+"
)


def _redact(value):
    return _REDACT_RE.sub(r"\1=***", str(value))

# ── Метаданные провайдера ────────────────────────────────────────────────────

NAME = "callibri"
LABEL = "Callibri"
BASE_URL = "https://api.callibri.ru"

# Структура учётных данных для UI
CREDENTIAL_FIELDS = [
    {"key": "email", "label": "Email", "secret": False, "env": "CALLIBRI_EMAIL"},
    {"key": "token", "label": "Token", "secret": True, "env": "CALLIBRI_TOKEN"},
]

# Типы обращений
APPEAL_TYPES = ["calls", "feedbacks", "chats", "emails"]

# Подписи типов для UI (ключ → отображение)
TYPE_LABELS = {
    "calls": "Звонки",
    "feedbacks": "Заявки",
    "chats": "Чаты",
    "emails": "Email",
}

# Колонки по умолчанию
DEFAULT_COLUMNS = [
    "date", "name_channel", "comment", "status", "type",
    "conversations_number", "utm_campaign",
]

# Максимум попыток на один чанк
MAX_RETRIES = 3

# Описания полей (для UI)
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

ALL_FIELDS = list(FIELD_DESCRIPTIONS.keys())


# ── Учётные данные ───────────────────────────────────────────────────────────

def check_credentials(creds):
    """Проверяем учётные данные. Возвращает (ok, message)."""
    email = (creds or {}).get("email", "")
    token = (creds or {}).get("token", "")
    if not email or not token:
        return False, "Заполни email и token"
    if "example" in email:
        return False, "Замени example-email на реальный"
    return True, "OK"


def _auth_params(creds):
    """Общие параметры авторизации для всех запросов."""
    return {"user_email": creds["email"], "user_token": creds["token"]}


# ── Период / чанки ───────────────────────────────────────────────────────────

def split_period(date1, date2):
    """Разбиваем период на чанки по 7 дней включительно (лимит API Callibri)."""
    chunks = []
    current = date1
    while current <= date2:
        chunk_end = min(current + timedelta(days=6), date2)
        chunks.append((current, chunk_end))
        current = chunk_end + timedelta(days=1)
    return chunks


# ── Утилиты ──────────────────────────────────────────────────────────────────

def _format_date(iso_date_str):
    """ISO → dd.mm.yyyy HH:MM."""
    if not iso_date_str:
        return ""
    try:
        dt = datetime.strptime(iso_date_str[:19], "%Y-%m-%dT%H:%M:%S")
        return dt.strftime("%d.%m.%Y %H:%M")
    except (ValueError, TypeError):
        return str(iso_date_str)


def _extract_appeal_id(appeal):
    """appeal_id или clbvid как запасной вариант (для дедупликации)."""
    aid = appeal.get("appeal_id")
    if aid:
        return str(aid)
    clbvid = appeal.get("clbvid")
    return str(clbvid) if clbvid else ""


# ── API ──────────────────────────────────────────────────────────────────────

def list_sites(creds):
    """Список проектов из API. Возвращает list[{site_id, sitename, domains}]."""
    resp = requests.get(
        f"{BASE_URL}/get_sites", params=_auth_params(creds), timeout=30
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("sites", [])


def _get_statistics(site_id, date1_str, date2_str, creds):
    """Статистика за период (макс 7 дней)."""
    params = {
        **_auth_params(creds),
        "site_id": site_id,
        "date1": date1_str,
        "date2": date2_str,
    }
    resp = requests.get(
        f"{BASE_URL}/site_get_statistics", params=params, timeout=60
    )
    resp.raise_for_status()
    return resp.json()


def test_connection(creds):
    """Проверка подключения. Возвращает (ok, site_count, message)."""
    try:
        ok, msg = check_credentials(creds)
        if not ok:
            return False, 0, msg
        sites = list_sites(creds)
        return True, len(sites), f"Подключено. Проектов: {len(sites)}"
    except Exception as e:
        return False, 0, f"Ошибка подключения: {e}"


def get_channels_and_statuses(site_id, creds):
    """Запрос статистики за 7 дней для извлечения имён каналов и статусов."""
    today = datetime.now()
    week_ago = today - timedelta(days=6)
    d1 = week_ago.strftime("%d.%m.%Y")
    d2 = today.strftime("%d.%m.%Y")
    data = _get_statistics(site_id, d1, d2, creds)

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


# ── Парсинг ──────────────────────────────────────────────────────────────────

def _build_row(appeal, ch_name, atype, columns):
    """Собираем строку из записи обращения."""
    row = {}
    for col in columns:
        if col == "date":
            row[col] = _format_date(appeal.get("date"))
        elif col == "name_channel":
            row[col] = ch_name
        elif col == "type":
            row[col] = atype
        else:
            row[col] = appeal.get(col, "") or ""
    return row


def _parse_chunk_data(data, channel_filter, seen_ids, columns,
                      type_filter=None, status_filter=None):
    """Парсим ответ чанка → {channel_name: [rows]}, кол-во записей."""
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
                appeal_id = _extract_appeal_id(appeal)
                if appeal_id and appeal_id in seen_ids:
                    continue
                if appeal_id:
                    seen_ids.add(appeal_id)
                if status_filter and (appeal.get("status", "") or "") not in status_filter:
                    continue
                channel_rows.append(_build_row(appeal, ch_name, atype, columns))

        if channel_rows:
            rows_by_channel[ch_name] = channel_rows
            chunk_count += len(channel_rows)

    return rows_by_channel, chunk_count


# ── Retry ────────────────────────────────────────────────────────────────────

def _emit(on_log, message):
    if on_log:
        on_log(message)
    else:
        log.info(message)


def _fetch_chunk_with_retry(site_id, d1, d2, idx, total_chunks, creds, on_log=None):
    """Запрос чанка с повторными попытками. Возвращает data или None.

    На 4xx (кроме 429) не ретраим — это клиентская ошибка, повтор не поможет.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return _get_statistics(site_id, d1, d2, creds)
        except requests.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status and 400 <= status < 500 and status != 429:
                _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): HTTP {status} — "
                              f"{_redact(e)} (клиентская ошибка, не ретраим)")
                return None
            _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): HTTP — {_redact(e)} "
                          f"(попытка {attempt}/{MAX_RETRIES})")
        except requests.RequestException as e:
            _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): сеть — {_redact(e)} "
                          f"(попытка {attempt}/{MAX_RETRIES})")

        if attempt < MAX_RETRIES:
            pause = attempt * 2
            _emit(on_log, f"    Повторная попытка через {pause}с...")
            time.sleep(pause)

    _emit(on_log, f"    Чанк {idx}/{total_chunks} ({d1}—{d2}): все {MAX_RETRIES} попытки исчерпаны — пропускаем")
    return None


# ── Обработка проекта ────────────────────────────────────────────────────────

def process_site(site, date1, date2, creds, filters=None, on_log=None, on_chunk=None):
    """
    Обработка одного проекта Callibri за весь период.

    Параметры:
      site — dict с site_id (из list_sites)
      date1, date2 — datetime границы периода
      creds — {"email", "token"}
      filters — dict с ключами:
        "channels": list[str] | None — фильтр по именам каналов
        "columns": list[str] | None — выбранные поля (default: DEFAULT_COLUMNS)
        "types": list[str] | None — фильтр по типам обращений
        "statuses": list[str] | None — фильтр по статусам
      on_log(msg) — callback логов
      on_chunk(cur, total) — callback прогресса чанков

    Возвращает (has_data: bool, rows_by_channel: dict, failed_chunks: int).
    """
    filters = filters or {}
    channel_filter = filters.get("channels")
    columns = filters.get("columns") or DEFAULT_COLUMNS
    type_filter = filters.get("types")
    status_filter = filters.get("statuses")

    site_id = site.get("site_id")
    chunks = split_period(date1, date2)
    total_chunks = len(chunks)

    merged = {}
    seen_ids = set()
    has_data = False
    failed_chunks = 0

    for idx, (chunk_start, chunk_end) in enumerate(chunks, start=1):
        d1 = chunk_start.strftime("%d.%m.%Y")
        d2 = chunk_end.strftime("%d.%m.%Y")

        data = _fetch_chunk_with_retry(site_id, d1, d2, idx, total_chunks, creds, on_log)
        if data is None:
            failed_chunks += 1
            if on_chunk:
                on_chunk(idx, total_chunks)
            continue

        chunk_rows, chunk_count = _parse_chunk_data(
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
