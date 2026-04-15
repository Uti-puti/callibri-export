# Callibri Export

GUI + CLI приложение для выгрузки обращений из API Callibri в XLSX/CSV.

## Возможности

- Выгрузка обращений (звонки, заявки, чаты, email) за произвольный период
- Автоматическая разбивка длинных периодов на 7-дневные чанки (лимит API)
- Настраиваемые поля выгрузки с описаниями для каждого проекта
- Фильтрация по каналам, типам обращений, статусам
- Формат XLSX или CSV
- Дедупликация, сортировка по дате, retry при ошибках API
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
```

### .exe
Скачай `CallibriExport.exe` из [Releases](../../releases), положи рядом `projects.json` и запусти.

## Настройка

1. Скопируй `projects.example.json` в `projects.json`
2. Заполни `site_id` и `folder` для своих проектов (список id — через GUI кнопка «+ Добавить» или `python explore.py`)
3. Укажи email и token из личного кабинета Callibri (в GUI или в файле `.env`)

## Структура

| Файл | Назначение |
|------|-----------|
| `app.py` | GUI-приложение (CustomTkinter) |
| `core.py` | Бизнес-логика (API, парсинг, запись файлов) |
| `export.py` | CLI-обёртка |
| `explore.py` | Разведка API (список проектов и полей) |
| `projects.json` | Конфигурация проектов (не в git) |
| `projects.example.json` | Пример конфигурации |
| `.env` | Email + token (не в git) |
| `FIELDS.md` | Справочник всех полей API |

## Сборка .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name CallibriExport app.py
```

Результат: `dist/CallibriExport.exe`
