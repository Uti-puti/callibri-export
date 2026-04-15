# Справочник полей для выгрузки

Поля указываются в `projects.json` в параметре `"fields"`.
Если `"fields"` не указан — используется набор по умолчанию.

## Набор по умолчанию

```json
"fields": ["date", "name_channel", "comment", "status", "type", "conversations_number", "utm_campaign"]
```

## Специальные (вычисляемые) поля

| Поле | Описание |
|------|----------|
| `date` | Дата обращения, формат `dd.mm.yyyy HH:MM` (конвертируется из ISO) |
| `name_channel` | Название канала (берётся из канала, не из записи) |
| `type` | Тип обращения: `calls`, `feedbacks`, `chats`, `emails` |

## Поля из API

Все остальные поля берутся напрямую из ответа API по имени ключа.
Актуальный список можно обновить через `python explore.py`.

| Поле | Описание | Пример значения |
|------|----------|-----------------|
| `appeal_id` | ID обращения (уникальный) | `71541938` |
| `phone` | Телефон клиента | `79226996022` |
| `email` | E-mail клиента | `nk@itcomms.ru` |
| `name` | Имя клиента | `Николай Косарев` |
| `comment` | Комментарий к обращению | `КТП` |
| `content` | Полный текст обращения (заявки/email) | `Добрый день, интересует...` |
| `status` | Статус обращения | `Лид`, `Нет ответа` |
| `source` | Источник трафика | `Google` |
| `traffic_type` | Тип трафика | `Переходы из поисковых систем` |
| `region` | Регион клиента | `Челябинская обл.` |
| `device` | Устройство | `desktop` |
| `conversations_number` | Порядковый номер обращения от клиента (1 = первое) | `1` |
| `is_lid` | Является ли лидом | `True` / `False` |
| `name_type` | Тип обращения (человекочитаемый) | `Звонок`, `Заявка`, `E-mail` |
| `landing_page` | Страница входа | `https://www.ec74.ru/...` |
| `lid_landing` | Целевая страница лида | `https://www.ec74.ru/catalogue.html` |
| `site_referrer` | Реферер | `google.com` |
| `link_download` | Ссылка на запись звонка | `https://api.callibri.ru/listens/...` |
| `duration` | Длительность звонка (сек) | `38` |
| `billsec` | Длительность разговора (сек) | `6` |
| `responsible_manager` | Ответственный менеджер | `Никто` |
| `responsible_manager_email` | Email менеджера | — |
| `call_status` | Статус звонка | — |
| `accurately` | Точное определение источника | `True` / `False` |
| `form_name` | Название формы (для заявок) | — |
| `utm_source` | UTM source | — |
| `utm_medium` | UTM medium | — |
| `utm_campaign` | UTM campaign | — |
| `utm_content` | UTM content | — |
| `utm_term` | UTM term | — |
| `query` | Поисковый запрос | — |
| `channel_id` | ID канала | `70200` |
| `crm_client_id` | ID клиента в CRM | `38817414` |
| `ym_uid` | Яндекс.Метрика UID | `1483485225` |
| `metrika_client_id` | Яндекс.Метрика Client ID | `1775109228651162734` |
| `ua_client_id` | Google Analytics Client ID | `1348104959.1775656133` |
| `clbvid` | Внутренний ID Callibri (запасной ключ) | `69d65cc4539160fc70ea7a06` |

## Пример настройки в projects.json

```json
{
  "site_id": 93946,
  "folder": "sparkcell",
  "channels": ["Контекст"],
  "fields": ["date", "name_channel", "name", "phone", "comment", "status", "type", "source", "utm_campaign", "landing_page"],
  "enabled": true
}
```

Не все поля заполняются для всех типов обращений — например, `duration` и `billsec` есть только у звонков, `form_name` — только у заявок.
