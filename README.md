# Callibri / Calltouch Export

GUI + CLI приложение для выгрузки обращений (звонки, заявки, чаты, email-обращения) из **Callibri** и **Calltouch** в XLSX / CSV / Google Sheets.

Написано на Python + CustomTkinter. Собирается в одиночный `.exe` (~31 МБ) — работает без установки Python.

---

## Возможности

### Источники данных
- **Два провайдера в одном приложении** — Callibri и Calltouch (выбор per-project)
- **Несколько аккаунтов Calltouch** с разными API-ключами — выгрузка сайтов с правами разных пользователей одной сборкой
- Выгрузка за произвольный период с автоматическим разбиением на чанки (7 дней у Callibri, 90 — у Calltouch) и пагинацией
- **Журнал звонков** + **Журнал заявок** + чаты (Calltouch); звонки / заявки / чаты / email (Callibri)
- Retry с уважением `Retry-After` на 429, понятные сообщения на 401/403

### Настройка выгрузки
- **Профили выгрузки**: один и тот же `site_id` можно добавить несколько раз с разными настройками (например, отдельные отчёты по cpc и organic трафику)
- Фильтры:
  - **Каналы / источники** (utmSource)
  - **Тип трафика** — utm_medium (cpc, organic, referral, social, email, прямые)
  - **Типы обращений** (calls / requests / chats / feedbacks / emails)
  - **Статусы** (текстовый список)
- **Выбираемые поля** с описаниями — отдельный набор для каждого провайдера, перетаскивание между «доступные ↔ выбранные», сортировка
- Дедупликация и сортировка по дате
- Independent чекбоксы у каждого профиля — состояние сохраняется в файле, не слетает при перерисовках UI

### Форматы вывода
- **XLSX** — отформатированный заголовок, автоширина колонок, защита от Formula Injection
- **CSV** — разделитель `;`, UTF-8 BOM
- **Google Sheets** — режим `append` (дополнить) или `replace` (заменить); автообрезка ячеек длиннее 50000 символов

### Безопасность
- Маскирование токенов в логах и сообщениях об ошибках (`clientApiId`, `user_token`, `apiKey` и т.п.)
- Защита от CSV/XLSX Formula Injection в значениях из API (комментариях клиентов и т.д.)
- Авторизация Google Sheets через Service Account

---

## Быстрый старт

### Использовать готовый .exe

1. Скачай `CallibriExport.exe` из [Releases](https://github.com/Uti-puti/callibri-export/releases)
2. Положи рядом — будут созданы автоматически:
   - `projects.json` (создастся при добавлении первого проекта через GUI)
   - `accounts.json` (создастся при добавлении первого аккаунта Calltouch)
   - `.env` (создастся при сохранении учётных данных)
   - `output/` (создастся при первом экспорте)
3. Запусти и заполни:
   - Учётные данные Callibri (email + token) или Calltouch (clientApiId)
   - Период
   - Добавь проекты («+ Добавить»)
   - Жми «Экспорт»

### Запуск из исходников

```bash
pip install -r callibri-export/requirements.txt

# GUI
python callibri-export/app.py

# CLI — последние 7 дней
python callibri-export/export.py

# CLI — за период
python callibri-export/export.py --date1 01.03.2026 --date2 06.04.2026

# CLI — N дней
python callibri-export/export.py --days 30

# Разведка API
python callibri-export/explore.py                        # Callibri
python callibri-export/explore.py --provider calltouch   # Calltouch
```

---

## Архитектура

Провайдер-агностик-ядро + пакет провайдеров.

```
callibri-export/
├── app.py                 # GUI на CustomTkinter
├── core.py                # Ядро: период, конфиг, запись XLSX/CSV, оркестрация run_export
├── accounts.py            # Управление именованными аккаунтами провайдеров
├── export.py              # CLI-обёртка
├── explore.py             # Разведка API (список сайтов + пример полей)
├── gsheets.py             # Google Sheets (gspread + service account)
├── providers/
│   ├── __init__.py        # get_provider / provider_names / all_providers
│   ├── callibri.py        # Callibri API: user_email+user_token, 7-дневные чанки
│   └── calltouch.py       # Calltouch API: clientApiId, журнал звонков/заявок
├── projects.json          # Конфигурация проектов (не в git)
├── accounts.json          # Аккаунты Calltouch с разными ключами (не в git)
├── .env                   # Учётные данные провайдеров (не в git)
└── credentials.json       # Service Account для Google Sheets (не в git)
```

Каждый провайдер реализует унифицированный интерфейс: `check_credentials`, `test_connection`, `list_sites`, `get_channels_and_statuses`, `process_site`. Это позволяет добавлять новых провайдеров без изменений в `core.py`/`app.py`.

---

## Конфигурация

### projects.json

```json
[
  {
    "provider": "calltouch",
    "site_id": 77163,
    "folder": "d-okna-cpc",
    "account": "d-okna",
    "types": ["calls", "requests"],
    "mediums": ["cpc"],
    "channels": ["yandex"],
    "fields": ["date", "type", "phone_number", "source", "medium", "utm_campaign", "status", "tags"],
    "format": "xlsx",
    "split_by_channel": false,
    "enabled": true,
    "gsheet": {
      "enabled": true,
      "spreadsheet_id": "1abc...",
      "sheet_name": "Лист1",
      "mode": "append"
    }
  }
]
```

| Ключ | Тип | Описание |
|---|---|---|
| `provider` | str | `"callibri"` (по умолчанию) или `"calltouch"` |
| `site_id` | int | ID сайта в системе провайдера |
| `folder` | str | Подпапка в `output/` (должна быть уникальной) |
| `account` | str | Имя аккаунта из `accounts.json` (только Calltouch). Если не указан — используется токен из `.env` |
| `types` | list | Типы обращений. Callibri: `calls/feedbacks/chats/emails`. Calltouch: `calls/requests/chats` |
| `mediums` | list | Фильтр utm_medium (Calltouch): `cpc, organic, referral, social, email, (none)` |
| `channels` | list | Фильтр по utmSource |
| `statuses` | list | Фильтр по статусу обращения |
| `fields` | list | Список и порядок колонок |
| `format` | str | `"xlsx"` (по умолчанию) или `"csv"` |
| `split_by_channel` | bool | Если true — каждый канал в отдельный файл |
| `file_export` | bool | Записывать ли локальный файл (по умолчанию true) |
| `enabled` | bool | Чекбокс в GUI пишет это значение |
| `gsheet` | obj | Конфигурация Google Sheets для этого проекта |

### accounts.json (Calltouch)

```json
{
  "calltouch": [
    {"name": "Kazna",  "client_api_id": "..."},
    {"name": "d-okna", "client_api_id": "..."}
  ]
}
```

Управление через GUI: кнопка «Аккаунты Calltouch…» на главном экране.

### .env

```ini
CALLIBRI_EMAIL=user@example.com
CALLIBRI_TOKEN=...
CALLTOUCH_API_ID=...
GSHEET_CREDENTIALS=credentials.json
```

---

## Google Sheets (опционально)

1. Создай проект в [Google Cloud Console](https://console.cloud.google.com/) → включи **Google Sheets API**
2. Создай **Service Account** → JSON-ключ → переименуй в `credentials.json` и положи рядом с приложением
3. Открой нужную Google-таблицу → «Поделиться» → добавь email сервисного аккаунта (есть в JSON-ключе)
4. В GUI: укажи путь к `credentials.json` → «Проверить»
5. В настройках проекта → вкладка «Google Sheets»: укажи таблицу (URL/ID), лист, режим (append/replace)

Ячейки длиннее 50000 символов автоматически усекаются с маркером `…[обрезано]` (лимит Google Sheets API).

---

## Сборка .exe

```bash
cd callibri-export
pip install pyinstaller
pyinstaller CallibriExport.spec --noconfirm
# Результат: dist/CallibriExport.exe (~31 МБ)
```

Готовые сборки публикуются в [Releases](https://github.com/Uti-puti/callibri-export/releases).

---

## API quirks

### Callibri
- Авторизация: query-параметры `user_email` + `user_token`
- `site_get_statistics`: максимум 7 дней включительно за один запрос
- `domains` в `get_sites` — строка, не список
- Дата в ISO: `2026-03-31T04:44:44.000Z`

### Calltouch
- Авторизация: query-параметр `clientApiId` (один токен на аккаунт; в этом приложении — несколько именованных аккаунтов через `accounts.json`)
- **Журнал звонков**: `/calls-service/RestAPI/{siteId}/calls-diary/calls`, формат даты `dd/mm/yyyy`, ответ `{records: [...]}`
- **Журнал заявок** (отдельный!): `/calls-service/RestAPI/requests?siteId=...`, формат даты `mm/dd/yyyy`, ответ — голый массив
- **Журнал сделок CRM**: `/orders-diary/orders` — содержит только заявки, попавшие в CRM-журнал через интеграцию (например amoCRM)
- Пагинация: `page` (1-based) + `limit` (до 1000)
- Тэги: поле `names` (список), а не `name`
- 4xx (кроме 429) не ретраятся; 401/403 даёт отдельное сообщение «токен отклонён» или «нет прав на сайт»

---

## Безопасность

- Все токены/email маскируются в логах и сообщениях об ошибках
- Защита от CSV/XLSX Formula Injection (`'` перед `= + - @ \t \r` в значениях)
- Google Sheets — только через Service Account (никаких OAuth с правами пользователя)

---

## Лицензия

MIT
