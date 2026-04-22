# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running scripts

```bash
# Установка зависимостей
C:/Users/5417268/AppData/Local/Python/pythoncore-3.14-64/python.exe -m pip install -r requirements.txt

# GUI-приложение
python app.py

# CLI — последние 7 дней (по умолчанию)
python export.py

# CLI — за произвольный период
python export.py --date1 01.03.2026 --date2 06.04.2026

# CLI — за последние N дней
python export.py --days 30

# Разведка API — проверка подключения, точные имена каналов и полей
python explore.py

# Сборка .exe
pip install pyinstaller
pyinstaller --onefile --windowed --name CallibriExport app.py
# Результат: dist/CallibriExport.exe (~16 МБ)
```

## Architecture

Провайдер-агностик-ядро + пакет провайдеров:

- **core.py** — общее ядро без привязки к API. Период, загрузка/сохранение `projects.json`, запись XLSX/CSV, отправка в Google Sheets, оркестрация `run_export()`. Делегирует выборку данных провайдеру по ключу `"provider"` в projects.json (дефолт — `callibri`). Ключевые функции:
  - `run_export(credentials, ...)` — принимает dict `{provider_name: {field: value}}` и диспатчит проекты
  - `get_project_provider(proj)` — вернуть модуль-провайдер по конфигу проекта
  - `resolve_period()`, `parse_date()`, `sanitize_filename()`, `write_xlsx()`, `write_csv()`, `load_projects()`, `save_projects()`
- **providers/** — пакет провайдеров. Каждый реализует единый интерфейс:
  - `NAME`, `LABEL`, `CREDENTIAL_FIELDS`, `FIELD_DESCRIPTIONS`, `ALL_FIELDS`, `DEFAULT_COLUMNS`, `TYPE_LABELS`
  - `check_credentials(creds)`, `test_connection(creds)`, `list_sites(creds)`
  - `get_channels_and_statuses(site_id, creds)`
  - `process_site(site, date1, date2, creds, filters, on_log, on_chunk)` → `(has_data, rows_by_channel, failed_chunks)`
  - `providers.get_provider(name)` / `providers.provider_names()` / `providers.all_providers()`
- **providers/callibri.py** — Callibri API. Авторизация query-параметрами `user_email`/`user_token`, чанки по 7 дней, структура `channels_statistics`.
- **providers/calltouch.py** — Calltouch API. Авторизация query-параметром `clientApiId`, эндпоинты `/calls-service/RestAPI/{siteId}/calls-diary/calls` и `/requests`, пагинация `page+limit`, период до 90 дней за чанк. «Каналы» ≈ `source`. Дедупликация по `(type, id)`.
- **export.py** — CLI. Собирает credentials из `.env` для всех провайдеров автоматически через `CREDENTIAL_FIELDS`.
- **explore.py** — разведка. Флаг `--provider callibri|calltouch`, показывает список сайтов и пример полей первой записи.
- **app.py** — GUI на CustomTkinter. Классы:
  - `App` — главное окно: поля учётных данных Callibri (email+token) и Calltouch (clientApiId), проверка подключения для каждого провайдера, период, проекты, прогресс, лог
  - `ProjectSettingsDialog` — настройки проекта с подстановкой полей/фильтров провайдера (`self.provider.FIELD_DESCRIPTIONS`, `TYPE_LABELS` и т.д.)
  - `AddProjectDialog` — добавление проекта (принимает `provider_name`, проставляет его в конфиг)
  - `ProviderChoiceDialog` — выбор провайдера перед добавлением проекта
- **gsheets.py** — автономный модуль Google Sheets (без изменений). Service Account авторизация через `credentials.json`. Ленивый импорт `gspread`/`google-auth`.

## GUI (app.py) — функционал

- CustomTkinter, тема system (тёмная/светлая по настройкам Windows)
- Экспорт в отдельном потоке (`threading.Thread` + `queue.Queue` + `root.after(100ms)`) — GUI не зависает
- **Настройки проекта** (кнопка «Настроить»):
  - Вкладка «Поля»: два списка (доступные ↔ выбранные) с описаниями из `FIELD_DESCRIPTIONS`, кнопки перемещения и сортировки, подсказка при клике на поле
  - Вкладка «Фильтры»: типы (чекбоксы), каналы (загрузка из API), статусы (текстовое поле), формат (XLSX/CSV), split by channel
- **Добавление проекта** (кнопка «+ Добавить»): загружает все проекты из API, автозаполнение folder
- **Удаление проекта** (кнопка «✕»)
- Все настройки сохраняются в `projects.json` через `core.save_projects()`
- Чекбоксы проектов переопределяют `enabled` из `projects.json` без изменения файла
- `.env` сохраняется автоматически при экспорте

## projects.json — конфигурация выгрузки

```json
{
  "provider": "callibri",
  "site_id": 4112,
  "folder": "energocenter",
  "channels": ["Канал 1", "Канал 2"],
  "split_by_channel": true,
  "enabled": true
}
```

- `provider` — опционально (`"callibri"` по умолчанию для обратной совместимости). Допустимые значения: `"callibri"`, `"calltouch"`.
- `site_id` + `folder` — обязательны. Список доступных id — из GUI («+ Добавить») или `explore.py --provider <name>`.
- `channels` — опционально. Фильтр по именам каналов. Настраивается в GUI (вкладка «Фильтры» → «Загрузить каналы из API»).
- `split_by_channel` — опционально (`false`). Если `true` — каждый канал в отдельный файл.
- `enabled` — опционально (`true`). Если `false` — проект пропускается. В GUI — снятый чекбокс.
- `fields` — опционально. Список полей и их порядок. Настраивается в GUI (вкладка «Поля»). По умолчанию: `["date", "name_channel", "comment", "status", "type", "conversations_number", "utm_campaign"]`.
- `types` — опционально. Фильтр по типу: `["calls", "feedbacks", "chats", "emails"]`. Настраивается в GUI.
- `statuses` — опционально. Фильтр по статусу: `["Лид", "Целевой"]`. Настраивается в GUI.
- `format` — опционально (`"xlsx"`). Формат: `"xlsx"` или `"csv"`. Настраивается в GUI.
- `gsheet` — опционально. Блок настроек Google Sheets: `enabled`, `spreadsheet_id`, `sheet_name`, `mode` (`"append"` / `"replace"`). Настраивается в GUI (вкладка «Google Sheets»).

## API quirks

### Callibri (providers/callibri.py)
- Авторизация — query-параметры `user_email` + `user_token` (не заголовки)
- `site_get_statistics`: максимальный период — **7 дней включительно** (оба конца). `timedelta(days=6)` = 7-дневный период. Для больших периодов `split_period()` бьёт на чанки
- `domains` в ответе `get_sites` — строка, не список
- `appeal_id` есть у всех типов обращений; `clbvid` — запасной ключ для дедупликации
- 429 при частых запросах → пауза `time.sleep(0.5)` между чанками и проектами
- Дата обращения в ISO формате: `2026-03-31T04:44:44.000Z`
- Retry: до 3 попыток на чанк, exponential backoff (2с, 4с). **4xx кроме 429 не ретраятся** (клиентская ошибка — повтор не поможет)

### Calltouch (providers/calltouch.py)
- Авторизация — query-параметр `clientApiId` (один токен на аккаунт)
- Базовый URL: `https://api.calltouch.ru`
- Пагинация: `page` (1-based) + `limit` (до 1000). Ответ содержит `records[]` + `recordsCount`/`pageTotal`
- Формат даты в запросе: `dd/mm/yyyy` (слеши!)
- Макс период за один чанк: **90 дней**. `split_period()` бьёт длинные периоды
- 429 — уважает заголовок `Retry-After`, иначе exponential backoff. **4xx кроме 429 не ретраятся** (401/403 → отдельное сообщение «токен отклонён»)
- Звонки и заявки — отдельные эндпоинты, мержатся в `rows_by_channel` по ключу `source`
- Дедупликация по `(type, id)`: id между сущностями могут пересекаться
- Эндпоинты могут потребовать уточнения при первом реальном запросе — провайдер толерантен к вариациям формата ответа (list vs `records` vs `items` vs `data`)
- `failed_chunks` инкрементируется только если **все эндпоинты чанка** упали; `has_data` зависит от фактического количества записей. В логе чанка — пометка «(сбой: N из M эндпоинтов)»

## Security

- **Redaction URL перед логированием** — оба провайдера содержат `_redact()`, маскирующий query-параметры `clientApiId`, `user_token`, `user_email`, `token`, `apiKey` перед вставкой URL в `ConnectionError`, `_emit` и любые другие пользовательские сообщения. Правило: любой `final_url` или `str(exception)` пропускается через `_redact()` **до** попадания в лог/ошибку.
- **CSV/XLSX formula injection guard** — `core._csv_safe()` префиксует `'` к значениям, начинающимся с `= + - @ \t \r`. Применяется автоматически в `write_csv` и `write_xlsx`. Без него значение из API (например, комментарий клиента) могло бы выполниться как формула в Excel у получателя.

## Output

Формат XLSX (openpyxl) с форматированной шапкой и автошириной колонок, или CSV (разделитель `;`, UTF-8 BOM). Выбирается через GUI или `"format"` в `projects.json`.

Колонки по умолчанию: `date`, `name_channel`, `comment`, `status`, `type`, `conversations_number`, `utm_campaign`. Настраиваются через GUI или `"fields"`.

Дедупликация по `appeal_id` внутри каждого проекта (общая между чанками одного проекта). Сортировка по дате.

Опциональный экспорт в Google Sheets: после записи файла данные отправляются в указанную таблицу через `gsheets.py`. Режимы: `append` (дополнить) / `replace` (заменить). Настраивается per-project через блок `"gsheet"` в `projects.json` и вкладку «Google Sheets» в GUI. Ошибка Sheets не блокирует файловый экспорт.

## Google Sheets (gsheets.py)

Автономный модуль. Service Account авторизация через `credentials.json`. Ленивый импорт — приложение работает без `gspread`/`google-auth`, если GSheets не настроен.

Конфигурация per-project в `projects.json`:
```json
"gsheet": {
  "enabled": true,
  "spreadsheet_id": "...",
  "sheet_name": "Лист1",
  "mode": "append"
}
```

Глобальный путь к credentials в `.env`: `GSHEET_CREDENTIALS=credentials.json`

Режим `append`: проверяет заголовок в строке 1, находит последнюю строку, дописывает. Предупреждает если заголовок не совпадает.
Режим `replace`: очищает лист, пишет заголовок + данные.

Запись пакетами по 500 строк (Google Sheets API квота: 60 req/min).

## Сборка .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name CallibriExport app.py
```

Результат: `dist/CallibriExport.exe` (~16 МБ). Рядом с `.exe` должен лежать `projects.json` (или создать через GUI). `.env` создаётся автоматически при первом экспорте. Папка `output/` создаётся автоматически.

Артефакты сборки (`build/`, `dist/`, `*.spec`) не коммитятся — добавлены в `.gitignore`.

## Windows console encoding

```python
sys.stdout.reconfigure(encoding="utf-8")
# Для logging — StreamHandler с sys.stdout, не stderr (иначе кракозябры)
```
