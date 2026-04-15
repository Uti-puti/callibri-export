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

Три модуля + разведка:

- **core.py** — бизнес-логика. Все функции API, парсинга, записи файлов. Без CLI, без `sys.exit` — ошибки через исключения (`ValueError`, `FileNotFoundError`). Email/token передаются параметрами. Callbacks `on_log`, `on_progress` для GUI. Ключевые функции:
  - `run_export()` — оркестратор экспорта с callbacks
  - `test_connection()` — проверка подключения к API
  - `get_app_dir()` — рабочая директория (для .exe — рядом с exe)
  - `save_projects()` — сохранение `projects.json` из GUI
  - `get_channels_and_statuses()` — запрос API за 1 день для получения имён каналов и статусов
  - `FIELD_DESCRIPTIONS` — словарь всех 40 полей с описаниями
  - `ALL_FIELDS` — список всех доступных полей
- **export.py** — CLI-обёртка. Парсит `argparse`, загружает `.env`, вызывает `core.run_export()`. ~60 строк.
- **app.py** — GUI на CustomTkinter. Три класса:
  - `App` — главное окно: подключение, период, проекты, прогресс-бар, лог, кнопки
  - `ProjectSettingsDialog` — диалог настроек проекта (поля с описаниями + фильтры каналов/типов/статусов + формат)
  - `AddProjectDialog` — диалог добавления проекта из API
- **explore.py** — разведка. Выводит список всех проектов и все поля первой записи каждого типа обращений.

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
  "site_id": 4332,
  "folder": "energ",
  "channels": ["Канал 1", "Канал 2"],
  "split_by_channel": true,
  "enabled": true
}
```

- `site_id` + `folder` — обязательны. Список доступных id — из GUI («+ Добавить») или `explore.py`.
- `channels` — опционально. Фильтр по именам каналов. Настраивается в GUI (вкладка «Фильтры» → «Загрузить каналы из API»).
- `split_by_channel` — опционально (`false`). Если `true` — каждый канал в отдельный файл.
- `enabled` — опционально (`true`). Если `false` — проект пропускается. В GUI — снятый чекбокс.
- `fields` — опционально. Список полей и их порядок. Настраивается в GUI (вкладка «Поля»). По умолчанию: `["date", "name_channel", "comment", "status", "type", "conversations_number", "utm_campaign"]`.
- `types` — опционально. Фильтр по типу: `["calls", "feedbacks", "chats", "emails"]`. Настраивается в GUI.
- `statuses` — опционально. Фильтр по статусу: `["Лид", "Целевой"]`. Настраивается в GUI.
- `format` — опционально (`"xlsx"`). Формат: `"xlsx"` или `"csv"`. Настраивается в GUI.

## API quirks

- Авторизация — query-параметры `user_email` + `user_token` (не заголовки)
- `site_get_statistics`: максимальный период — **7 дней включительно** (оба конца). `timedelta(days=6)` = 7-дневный период. Для больших периодов скрипт автоматически разбивает на чанки через `split_period()`
- `domains` в ответе `get_sites` — строка, не список
- `appeal_id` есть у всех типов обращений; `clbvid` — запасной ключ для дедупликации
- API возвращает 429 при частых запросах → пауза `time.sleep(0.5)` между чанками и проектами
- Дата обращения в ISO формате: `2026-03-31T04:44:44.000Z`
- Retry: до 3 попыток на чанк, exponential backoff (2с, 4с)

## Output

Формат XLSX (openpyxl) с форматированной шапкой и автошириной колонок, или CSV (разделитель `;`, UTF-8 BOM). Выбирается через GUI или `"format"` в `projects.json`.

Колонки по умолчанию: `date`, `name_channel`, `comment`, `status`, `type`, `conversations_number`, `utm_campaign`. Настраиваются через GUI или `"fields"`.

Дедупликация по `appeal_id` внутри каждого проекта (общая между чанками одного проекта). Сортировка по дате.

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
