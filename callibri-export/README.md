# Appeal Export (Callibri + Calltouch)

GUI + CLI приложение для выгрузки обращений из API **Callibri** и **Calltouch** в XLSX/CSV.

## Возможности

- **Два провайдера в одном приложении**: Callibri и Calltouch (выбор per-project через `"provider"` в `projects.json`)
- Выгрузка обращений за произвольный период
- Автоматическая разбивка длинных периодов на чанки (7 дней у Callibri, 90 — у Calltouch) + пагинация
- Настраиваемые поля выгрузки с описаниями — отдельный набор для каждого провайдера
- Фильтрация по каналам/источникам, типам обращений, статусам
- Формат XLSX или CSV
- Экспорт в Google Sheets (append/replace) — опционально, через Service Account
- Дедупликация, сортировка по дате, retry при ошибках API (уважает `Retry-After`)
- GUI на CustomTkinter — настройки прямо в интерфейсе
- CLI для автоматизации и скриптов
- Сборка в `.exe` (PyInstaller) — работает без установки Python

## Быстрый старт

### GUI
```bash
pip install -r requirements.txt
python app.py
```

### CLI
```bash
python export.py --days 7
python export.py --date1 01.03.2026 --date2 06.04.2026

# Разведка API (список проектов + пример полей)
python explore.py                         # Callibri (по умолчанию)
python explore.py --provider calltouch    # Calltouch
```

### .exe
Скачай `CallibriExport.exe` из [Releases](../../releases), положи рядом `projects.json` и запусти.

## Настройка

1. Скопируй `projects.example.json` в `projects.json`
2. Для каждого проекта заполни `"provider"` (`"callibri"` или `"calltouch"`), `site_id` и `folder`. Списки id:
   - GUI: кнопка «+ Добавить» (спросит провайдера, если их > 1)
   - CLI: `python explore.py --provider callibri|calltouch`
3. Учётные данные в `.env` или GUI:
   - Callibri: `CALLIBRI_EMAIL` + `CALLIBRI_TOKEN`
   - Calltouch: `CALLTOUCH_API_ID` (clientApiId из личного кабинета)

## Структура

| Файл | Назначение |
|------|-----------|
| `app.py` | GUI-приложение (CustomTkinter) |
| `core.py` | Бизнес-логика (API, парсинг, запись файлов) |
| `export.py` | CLI-обёртка |
| `explore.py` | Разведка API (список проектов и полей) |
| `gsheets.py` | Модуль Google Sheets (авторизация, запись) |
| `projects.json` | Конфигурация проектов (не в git) |
| `projects.example.json` | Пример конфигурации |
| `credentials.json` | Ключ Service Account для Google Sheets (не в git) |
| `.env` | Email + token + путь к credentials (не в git) |
| `FIELDS.md` | Справочник всех полей API |

## Google Sheets (опционально)

1. Создай проект в [Google Cloud Console](https://console.cloud.google.com/), включи Google Sheets API
2. Создай Service Account, скачай JSON-ключ → `credentials.json` рядом с приложением
3. Открой нужную Google-таблицу → «Поделиться» → добавь email сервисного аккаунта (из JSON-ключа)
4. В GUI: укажи путь к `credentials.json` (секция «Google Sheets») → «Проверить»
5. В настройках проекта (кнопка «Настроить» → вкладка «Google Sheets»): укажи таблицу, лист, режим (append/replace)

## Сборка .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name CallibriExport app.py
```

Результат: `dist/CallibriExport.exe`
