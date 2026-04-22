"""
providers/calltouch.py — провайдер Calltouch API.

API-специфика:
- Базовый URL: https://api.calltouch.ru
- Авторизация: query-параметр clientApiId (один токен на аккаунт)
- Эндпоинт звонков:  GET /calls-service/RestAPI/{siteId}/calls-diary/calls
- Эндпоинт сделок (универсальный журнал):
                     GET /calls-service/RestAPI/{siteId}/orders-diary/orders
  Поддерживает фильтр orderSource=CALL|REQUEST|CHAT. Используется для
  заявок и чатов (для звонков оставлен calls-diary как более богатый).
  Документация: https://www.calltouch.ru/support/api-metod-vygruzki-zhurnala-sdelok/
- Эндпоинт сайтов:   GET /sites-service/sites (список доступных siteId аккаунта)
- Формат даты: dd/mm/yyyy (dateFrom / dateTo)
- Пагинация: page (>=1), limit (<=1000). Ответ содержит records[] + recordsCount + pageTotal
- Макс период за один запрос: 3 месяца

Доп. параметры для звонков (включены по умолчанию):
  withCallTags=true, withYandexDirect=true, withAttributionFields=true
Для orders-diary:
  withOrdersTags=true, withComments=true, withContacts=true

Credentials: {"client_api_id": str}

Объединяем calls, requests и chats в общий поток Row с типом.
Дедупликация по (type, id): id между сущностями могут пересекаться.
"""

import re
import time
import logging
from datetime import datetime, timedelta

import requests

log = logging.getLogger(__name__)

# Маскируем секреты в query-параметрах перед попаданием URL в логи/ошибки.
_SECRET_QUERY_KEYS = ("clientApiId", "user_token", "token", "apiKey", "api_key")
_REDACT_RE = re.compile(
    r"(?i)(" + "|".join(_SECRET_QUERY_KEYS) + r")=[^&\s]+"
)


def _redact(value):
    return _REDACT_RE.sub(r"\1=***", str(value))

# ── Метаданные ────────────────────────────────────────────────────────────────

NAME = "calltouch"
LABEL = "Calltouch"
BASE_URL = "https://api.calltouch.ru"

CREDENTIAL_FIELDS = [
    {"key": "client_api_id", "label": "clientApiId", "secret": True,
     "env": "CALLTOUCH_API_ID"},
]

# Типы записей (в терминах Calltouch).
# По умолчанию включены все три: звонки, заявки, чаты.
APPEAL_TYPES = ["calls", "requests", "chats"]

TYPE_LABELS = {
    "calls": "Звонки",
    "requests": "Заявки",
    "chats": "Чаты",
}

# Колонки по умолчанию
DEFAULT_COLUMNS = [
    "date", "type", "phone_number", "source", "medium",
    "utm_campaign", "status", "duration",
]

# Максимум попыток на один запрос
MAX_RETRIES = 3

# Максимальный период за один запрос — 3 месяца. Длиннее — разбиваем.
MAX_PERIOD_DAYS = 90

# Размер страницы (Calltouch допускает до 1000)
PAGE_LIMIT = 1000

# Описания полей (поля унифицированы; недостающие в конкретном типе остаются пустыми)
FIELD_DESCRIPTIONS = {
    # Специальные
    "date": "Дата (dd.mm.yyyy HH:MM)",
    "type": "Тип записи (calls / requests / chats)",
    # Общие
    "id": "ID записи (callId / orderId)",
    "phone_number": "Телефон клиента",
    "client_name": "Имя клиента",
    "client_email": "Email клиента",
    "source": "Источник (utm_source / канал Calltouch)",
    "medium": "Канал (utm_medium)",
    "utm_campaign": "UTM campaign",
    "utm_content": "UTM content",
    "utm_term": "UTM term",
    "keyword": "Поисковый запрос",
    "city": "Город клиента",
    "ref": "Реферер",
    "url": "Страница входа",
    "status": "Статус звонка / обработки заявки",
    "tags": "Теги (через запятую)",
    "attribution": "Модель атрибуции",
    "unique_call": "Уникальный звонок (True/False)",
    "target_call": "Целевой звонок (True/False)",
    "ya_client_id": "Яндекс.Метрика ClientID",
    "ga_client_id": "Google Analytics ClientID",
    "site_id": "ID проекта Calltouch",
    # Только звонки
    "duration": "Длительность звонка, сек",
    "waiting_connect": "Время ожидания соединения, сек",
    "call_url": "Ссылка на запись звонка",
    "manager": "Ответственный менеджер",
    # Только заявки / чаты
    "form_name": "Название формы",
    "comment": "Комментарий / содержимое заявки",
    "order_number": "Номер заявки (orderNumber)",
    "order_name": "Название сделки (orderName)",
    "order_status": "Статус сделки (orderStatus)",
    "request_type": "Тип источника (CALL / REQUEST / CHAT)",
    "session_id": "Идентификатор сессии",
    "created_date": "Дата создания сделки",
    "updated_date": "Дата обновления сделки",
    "planned_amount": "Планируемая сумма",
    "completed_amount": "Завершённая сумма",
    "funnel": "Воронка",
    "service": "Сервис / продукт",
}

ALL_FIELDS = list(FIELD_DESCRIPTIONS.keys())


# ── Учётные данные ───────────────────────────────────────────────────────────

def check_credentials(creds):
    token = (creds or {}).get("client_api_id", "")
    if not token:
        return False, "Укажи clientApiId (токен Calltouch)"
    if "example" in token.lower():
        return False, "Замени example-токен на реальный"
    return True, "OK"


def _auth_params(creds):
    return {"clientApiId": creds["client_api_id"]}


# ── Разбиение периода ────────────────────────────────────────────────────────

def split_period(date1, date2):
    """Разбиваем период на чанки по MAX_PERIOD_DAYS дней включительно."""
    chunks = []
    current = date1
    while current <= date2:
        chunk_end = min(current + timedelta(days=MAX_PERIOD_DAYS - 1), date2)
        chunks.append((current, chunk_end))
        current = chunk_end + timedelta(days=1)
    return chunks


# ── Утилиты ──────────────────────────────────────────────────────────────────

def _emit(on_log, message):
    if on_log:
        on_log(message)
    else:
        log.info(message)


def _format_date(iso_date_str):
    """ISO / dd/mm/yyyy HH:mm:ss → dd.mm.yyyy HH:MM."""
    if not iso_date_str:
        return ""
    value = str(iso_date_str).strip()
    formats = [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(value[:19], fmt)
            return dt.strftime("%d.%m.%Y %H:%M")
        except ValueError:
            continue
    return value


def _record_date_key(record):
    """Извлечь сырую дату для сортировки (строка)."""
    for key in ("callTime", "date", "requestDate", "created"):
        v = record.get(key)
        if v:
            return str(v)
    return ""


# ── Извлечение полей ─────────────────────────────────────────────────────────

def _utm(record):
    """Приводит utm_* и attribution к плоскому виду (в ответе могут быть в attribution)."""
    attrs = record.get("attributionSources") or record.get("attributes") or {}
    if isinstance(attrs, list) and attrs:
        attrs = attrs[0] or {}
    if not isinstance(attrs, dict):
        attrs = {}

    def pick(*keys):
        for k in keys:
            if record.get(k):
                return record.get(k)
            if attrs.get(k):
                return attrs.get(k)
        return ""

    return {
        "source": pick("source", "utmSource", "utm_source"),
        "medium": pick("medium", "utmMedium", "utm_medium"),
        "utm_campaign": pick("utmCampaign", "utm_campaign", "campaign"),
        "utm_content": pick("utmContent", "utm_content"),
        "utm_term": pick("utmTerm", "utm_term"),
        "keyword": pick("keyword", "searchQuery"),
    }


def _build_row_calls(record, columns):
    """Строка для звонка."""
    utm = _utm(record)
    tags = record.get("tags") or record.get("callTags") or []
    if isinstance(tags, list):
        tag_names = [t.get("name") if isinstance(t, dict) else str(t) for t in tags]
        tags_str = ", ".join(t for t in tag_names if t)
    else:
        tags_str = str(tags) if tags else ""

    base = {
        "id": str(record.get("callId") or record.get("id") or ""),
        "date": _format_date(
            record.get("callTime") or record.get("date") or record.get("callDate")
        ),
        "type": "calls",
        "phone_number": (record.get("phoneNumber") or record.get("clientPhone")
                         or record.get("phone") or ""),
        "client_name": record.get("clientName") or record.get("name") or "",
        "client_email": record.get("clientEmail") or record.get("email") or "",
        **utm,
        "city": record.get("city") or record.get("callerCity") or "",
        "ref": record.get("ref") or record.get("referrer") or "",
        "url": record.get("url") or record.get("landingPage") or "",
        "status": record.get("callStatus") or record.get("status") or "",
        "tags": tags_str,
        "attribution": record.get("attribution") or "",
        "unique_call": record.get("uniqueCall", ""),
        "target_call": record.get("targetCall", ""),
        "ya_client_id": record.get("yaClientId") or "",
        "ga_client_id": record.get("gaClientId") or "",
        "site_id": record.get("siteId") or "",
        "duration": record.get("duration", ""),
        "waiting_connect": record.get("waitingConnect", ""),
        "call_url": record.get("callUrl") or record.get("recordUrl") or "",
        "manager": record.get("manager") or record.get("managerName") or "",
        "form_name": "",
        "comment": record.get("comment") or "",
    }
    return {col: base.get(col, "") for col in columns}


def _first(lst):
    """Первый непустой элемент списка или ''."""
    if isinstance(lst, list):
        for v in lst:
            if v:
                return v
    return ""


def _extract_order_source(record):
    """Источник сделки из orders-diary: visit.utmSource → source/utm_source.
    Если нет — использует upper-level поля, которые могут быть в упрощённой
    форме ответа.
    """
    visit = record.get("visit") or {}
    return (visit.get("utmSource") or visit.get("source")
            or record.get("source") or "")


def _build_row_orders(record, columns, row_type="requests"):
    """Строка из записи orders-diary (заявка / чат / сделка).

    Маппинг по документации Calltouch:
      https://www.calltouch.ru/support/api-metod-vygruzki-zhurnala-sdelok/
    Record shape:
      { orderId, orderNumber, orderName, orderStatus, createdDate, orderDate,
        client: {fio, phones[], emails[]},
        visit: {utmSource, utmMedium, utmCampaign, city, url, sessionId, ...},
        orderSource: {type: CALL|REQUEST|CHAT, formName, callId, duration, ...},
        tags: [...], comment, manager, ... }
    """
    client = record.get("client") or {}
    visit = record.get("visit") or {}
    source_obj = record.get("orderSource") or {}
    yandex_direct = visit.get("yandexDirect") or {}
    google_ads = visit.get("googleAdWords") or {}

    tags = record.get("tags") or []
    if isinstance(tags, list):
        tag_names = [t.get("name") if isinstance(t, dict) else str(t) for t in tags]
        tags_str = ", ".join(n for n in tag_names if n)
    else:
        tags_str = str(tags) if tags else ""

    # Телефон и email могут приходить и плоско, и внутри client
    phone = (_first(client.get("phones"))
             or record.get("phoneNumber") or record.get("phone") or "")
    if isinstance(phone, dict):
        phone = phone.get("value") or ""

    email = (_first(client.get("emails"))
             or record.get("email") or record.get("clientEmail") or "")
    if isinstance(email, dict):
        email = email.get("value") or ""

    fio = (client.get("fio") or record.get("fio")
           or record.get("clientName") or "")

    # Манагер может быть dict или строка
    manager = record.get("manager") or ""
    if isinstance(manager, dict):
        manager = manager.get("name") or manager.get("fio") or ""

    # Комментарий может быть массивом
    comment = record.get("comment")
    if isinstance(comment, list):
        parts = []
        for c in comment:
            if isinstance(c, dict):
                parts.append(c.get("text") or c.get("comment") or "")
            else:
                parts.append(str(c))
        comment = " | ".join(p for p in parts if p)

    base = {
        "id": str(record.get("orderId") or record.get("id") or ""),
        "date": _format_date(
            record.get("createdDate") or record.get("orderDate")
            or record.get("updatedDate")
        ),
        "type": row_type,
        "phone_number": str(phone) if phone else "",
        "client_name": fio,
        "client_email": str(email) if email else "",
        "source": visit.get("utmSource") or visit.get("source") or "",
        "medium": visit.get("utmMedium") or "",
        "utm_campaign": visit.get("utmCampaign") or "",
        "utm_content": visit.get("utmContent") or "",
        "utm_term": visit.get("utmTerm") or "",
        "keyword": (visit.get("keyword") or yandex_direct.get("keyword")
                    or google_ads.get("keyword") or ""),
        "city": visit.get("city") or "",
        "ref": visit.get("ref") or visit.get("referrer") or "",
        "url": visit.get("url") or "",
        "status": record.get("orderStatus") or "",
        "tags": tags_str,
        "attribution": visit.get("attribution") or "",
        "unique_call": "",
        "target_call": "",
        "ya_client_id": visit.get("yaClientId") or visit.get("ym_uid") or "",
        "ga_client_id": visit.get("gaClientId") or visit.get("ga_uid") or "",
        "site_id": record.get("siteId") or "",
        "duration": source_obj.get("duration") or "",
        "waiting_connect": "",
        "call_url": source_obj.get("callUrl") or source_obj.get("recordUrl") or "",
        "manager": str(manager),
        "form_name": source_obj.get("formName") or record.get("orderName") or "",
        "comment": str(comment) if comment else "",
        "order_number": str(record.get("orderNumber") or ""),
        "order_name": record.get("orderName") or "",
        "order_status": record.get("orderStatus") or "",
        "request_type": source_obj.get("type") or "",
        "session_id": visit.get("sessionId") or "",
        "created_date": _format_date(record.get("createdDate")),
        "updated_date": _format_date(record.get("updatedDate")),
        "planned_amount": record.get("plannedAmount") or "",
        "completed_amount": record.get("completedAmount") or "",
        "funnel": record.get("funnel") or "",
        "service": record.get("service") or "",
    }
    return {col: base.get(col, "") for col in columns}


# ── API-запросы ──────────────────────────────────────────────────────────────

_HEADERS = {
    "Accept": "application/json",
    "User-Agent": "CallibriExport/1.0 (+https://github.com/)",
}


def _request_with_retry(url, params, on_log=None, label=""):
    """GET с retry. Возвращает JSON или бросает ConnectionError с понятным сообщением.

    На 4xx (кроме 429) не ретраим — это клиентская ошибка (неверный токен,
    некорректный siteId и т.д.), повтор не поможет и лишь затянет UX.
    """
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, params=params, headers=_HEADERS, timeout=60,
                                allow_redirects=False)
            final_url = _redact(resp.url)
            if resp.status_code in (301, 302, 303, 307, 308):
                loc = resp.headers.get("Location", "?")
                raise ConnectionError(
                    f"эндпоинт {_redact(url)} недоступен (redirect → {loc}). "
                    f"Скорее всего этот API не поддерживается Calltouch "
                    f"через clientApiId."
                )
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", attempt * 2))
                _emit(on_log, f"    {label}: 429 Too Many Requests — пауза {retry_after}с")
                time.sleep(retry_after)
                continue
            if resp.status_code in (401, 403):
                raise ConnectionError(
                    f"HTTP {resp.status_code} на {final_url} — "
                    f"токен отклонён (проверь clientApiId)"
                )
            if 400 <= resp.status_code < 500:
                body = resp.text[:300]
                raise ConnectionError(
                    f"HTTP {resp.status_code} на {final_url} — {body}"
                )
            if not resp.ok:
                body = resp.text[:300]
                raise ConnectionError(
                    f"HTTP {resp.status_code} на {final_url} — {body}"
                )
            try:
                return resp.json()
            except ValueError:
                body = resp.text[:300].replace("\n", " ").replace("\r", "")
                ctype = resp.headers.get("Content-Type", "")
                raise ConnectionError(
                    f"ответ не JSON на {final_url} "
                    f"(Content-Type: {ctype}); body: {body!r}"
                )
        except requests.RequestException as e:
            last_exc = e
            _emit(on_log, f"    {label}: ошибка — {_redact(e)} (попытка {attempt}/{MAX_RETRIES})")
            if attempt < MAX_RETRIES:
                time.sleep(attempt * 2)
    raise last_exc or requests.RequestException(f"{label}: все попытки исчерпаны")


_SITES_ENDPOINTS = [
    "/sites-service/sites",
    "/calls-service/RestAPI/sites",
    "/calls-service/RestAPI/account/sites",
    "/account-api/v1/sites",
    "/account-api/v1/siteblocks",
]


def _normalize_sites(data):
    """Извлечь список сайтов из произвольной формы ответа Calltouch."""
    if isinstance(data, list):
        items = data
    elif isinstance(data, dict):
        items = (data.get("records") or data.get("sites")
                 or data.get("items") or data.get("data") or [])
    else:
        items = []

    sites = []
    for item in items:
        if not isinstance(item, dict):
            continue
        sid = item.get("siteId") or item.get("id") or item.get("site_id")
        name = (item.get("name") or item.get("siteName")
                or item.get("siteblockName") or str(sid))
        domains = (item.get("domain") or item.get("domainName")
                   or item.get("siteUrl") or "")
        if sid:
            sites.append({
                "site_id": sid,
                "sitename": name,
                "domains": domains,
            })
    return sites


def list_sites(creds):
    """Попытаться получить список сайтов через известные эндпоинты.

    Публичного API /sites у Calltouch нет — все попытки возвращают HTML
    справочного центра. Функция возвращает пустой список, если ничего не
    удалось распознать: вызывающий код должен предложить ручной ввод siteId.
    """
    for path in _SITES_ENDPOINTS:
        url = f"{BASE_URL}{path}"
        try:
            data = _request_with_retry(
                url, _auth_params(creds), label=f"list_sites {path}",
            )
        except (requests.RequestException, ConnectionError):
            continue

        sites = _normalize_sites(data)
        if sites:
            log.info("Calltouch: list_sites through %s — %d sites", path, len(sites))
            return sites
    return []


# Флаг: провайдер требует ручного ввода siteId (нет публичного /sites API)
REQUIRES_MANUAL_SITE_ID = True


def test_connection(creds):
    """У Calltouch нет публичного эндпоинта /sites. Проверка лишь валидирует
    наличие токена — реальная проверка токена случится при запросе данных."""
    ok, msg = check_credentials(creds)
    if not ok:
        return False, 0, msg
    # Попробуем всё же получить список сайтов (вдруг аккаунт имеет доступ)
    try:
        sites = list_sites(creds)
    except Exception:
        sites = []
    if sites:
        return True, len(sites), f"Подключено. Проектов: {len(sites)}"
    return True, 0, (
        "Токен задан. Публичного списка сайтов у Calltouch нет — "
        "siteId нужно указать вручную (ЛК → Интеграции → API)."
    )


def _fetch_paginated(endpoint_path, site_id, date1, date2, creds,
                     extra_params=None, on_log=None, label=""):
    """Генератор записей через пагинацию page+limit."""
    url = f"{BASE_URL}{endpoint_path}"
    params = {
        **_auth_params(creds),
        "dateFrom": date1.strftime("%d/%m/%Y"),
        "dateTo": date2.strftime("%d/%m/%Y"),
        "limit": PAGE_LIMIT,
    }
    if extra_params:
        params.update(extra_params)

    page = 1
    while True:
        params["page"] = page
        data = _request_with_retry(url, params, on_log, f"{label} page={page}")

        if isinstance(data, list):
            records = data
            has_more = len(records) == PAGE_LIMIT
        else:
            records = data.get("records") or data.get("items") or data.get("data") or []
            total = data.get("recordsCount") or data.get("total") or data.get("pageTotal") or 0
            has_more = page * PAGE_LIMIT < total if total else (len(records) == PAGE_LIMIT)

        for r in records:
            yield r

        if not records or not has_more:
            break

        page += 1
        time.sleep(0.3)


def get_channels_and_statuses(site_id, creds):
    """Для Calltouch «каналы» ≈ источники (source), статусы — callStatus / orderStatus.

    Берём выборку за последние 30 дней из calls-diary и orders-diary,
    собираем уникальные значения.
    """
    date2 = datetime.now()
    date1 = date2 - timedelta(days=29)

    sources = set()
    statuses = set()

    # Звонки
    try:
        for rec in _fetch_paginated(
            f"/calls-service/RestAPI/{site_id}/calls-diary/calls",
            site_id, date1, date2, creds,
            extra_params={"withAttributionFields": "true"},
            label="sample calls",
        ):
            src = rec.get("source") or (rec.get("attributionSources") or [{}])[0].get("source")
            if src:
                sources.add(str(src))
            st = rec.get("callStatus") or rec.get("status")
            if st:
                statuses.add(str(st))
    except (requests.RequestException, ConnectionError):
        pass

    # Сделки (заявки + чаты) — только источники/статусы, без фильтра по типу
    try:
        for rec in _fetch_paginated(
            f"/calls-service/RestAPI/{site_id}/orders-diary/orders",
            site_id, date1, date2, creds,
            extra_params={"withOrdersTags": "false"},
            label="sample orders",
        ):
            src = _extract_order_source(rec)
            if src:
                sources.add(str(src))
            st = rec.get("orderStatus")
            if st:
                statuses.add(str(st))
    except (requests.RequestException, ConnectionError):
        pass

    return sorted(sources), sorted(statuses)


# ── Обработка проекта ────────────────────────────────────────────────────────

def process_site(site, date1, date2, creds, filters=None, on_log=None, on_chunk=None):
    """
    Обработка одного проекта Calltouch за период.
    Возвращает (has_data, rows_by_channel, failed_chunks).

    «Каналы» здесь = группировка по source (для совместимости с GUI split_by_channel).
    """
    filters = filters or {}
    channel_filter = filters.get("channels")  # список source-значений
    columns = filters.get("columns") or DEFAULT_COLUMNS
    type_filter = filters.get("types") or APPEAL_TYPES
    status_filter = filters.get("statuses")

    site_id = site.get("site_id")
    chunks = split_period(date1, date2)
    total_chunks = len(chunks)

    rows_by_channel = {}
    seen_ids = set()
    failed_chunks = 0
    has_data = False

    # Доп. параметры для звонков — подтянуть теги, атрибуцию, данные Директа.
    call_extra_params = {
        "withCallTags": "true",
        "withAttributionFields": "true",
        "withYandexDirect": "true",
    }
    # Универсальные параметры для orders-diary.
    orders_base_params = {
        "withOrdersTags": "true",
        "withComments": "true",
        "withContacts": "true",
    }

    orders_path = f"/calls-service/RestAPI/{site_id}/orders-diary/orders"

    for idx, (chunk_start, chunk_end) in enumerate(chunks, start=1):
        d1 = chunk_start.strftime("%d.%m.%Y")
        d2 = chunk_end.strftime("%d.%m.%Y")
        chunk_count = 0
        endpoints_total = 0
        endpoints_ok = 0

        endpoints = []
        if "calls" in type_filter:
            endpoints.append((
                "calls",
                f"/calls-service/RestAPI/{site_id}/calls-diary/calls",
                _build_row_calls,
                call_extra_params,
                "calls",
            ))
        if "requests" in type_filter:
            endpoints.append((
                "requests",
                orders_path,
                lambda r, c: _build_row_orders(r, c, row_type="requests"),
                {**orders_base_params, "orderSource": "REQUEST"},
                "orders",
            ))
        if "chats" in type_filter:
            endpoints.append((
                "chats",
                orders_path,
                lambda r, c: _build_row_orders(r, c, row_type="chats"),
                {**orders_base_params, "orderSource": "CHAT"},
                "orders",
            ))

        for atype, path, builder, extra, shape in endpoints:
            endpoints_total += 1
            try:
                records_iter = list(_fetch_paginated(
                    path, site_id, chunk_start, chunk_end, creds,
                    extra_params=extra,
                    on_log=on_log,
                    label=f"{atype} чанк {idx}/{total_chunks}",
                ))
                endpoints_ok += 1
            except (requests.RequestException, ConnectionError) as e:
                _emit(on_log, f"    Чанк {idx}/{total_chunks} [{atype}] ({d1}—{d2}): пропущен — {_redact(e)}")
                continue

            for record in records_iter:
                if shape == "orders":
                    rid = str(record.get("orderId") or record.get("id") or "")
                else:
                    rid = str(record.get("callId") or record.get("id") or "")
                dedup_key = (atype, rid) if rid else None
                if dedup_key and dedup_key in seen_ids:
                    continue
                if dedup_key:
                    seen_ids.add(dedup_key)

                if shape == "orders":
                    status_val = record.get("orderStatus") or ""
                    src = _extract_order_source(record) or "Без источника"
                else:
                    status_val = record.get("callStatus") or record.get("status") or ""
                    src = (record.get("source")
                           or ((record.get("attributionSources") or [{}])[0] or {}).get("source")
                           or "Без источника")

                if status_filter and status_val not in status_filter:
                    continue
                if channel_filter and str(src) not in channel_filter:
                    continue

                row = builder(record, columns)
                rows_by_channel.setdefault(str(src), []).append(row)
                chunk_count += 1

        # Чанк считается провальным, если ни один эндпоинт не отработал.
        if endpoints_total > 0 and endpoints_ok == 0:
            failed_chunks += 1
        if chunk_count > 0:
            has_data = True

        summary = f"    Чанк {idx}/{total_chunks}: {d1} — {d2} — {chunk_count} записей"
        if endpoints_total and endpoints_ok < endpoints_total:
            summary += f" (сбой: {endpoints_total - endpoints_ok} из {endpoints_total} эндпоинтов)"
        _emit(on_log, summary)

        if on_chunk:
            on_chunk(idx, total_chunks)

        if idx < total_chunks:
            time.sleep(0.5)

    return has_data, rows_by_channel, failed_chunks
