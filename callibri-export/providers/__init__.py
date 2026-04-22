"""
providers/ — пакет с провайдерами API для выгрузки обращений.

Каждый провайдер — отдельный модуль с единым интерфейсом:
- NAME, LABEL — идентификатор и отображаемое имя
- CREDENTIAL_FIELDS — описание учётных данных для UI
- FIELD_DESCRIPTIONS, ALL_FIELDS, DEFAULT_COLUMNS — поля выгрузки
- check_credentials(creds), test_connection(creds), list_sites(creds)
- get_channels_and_statuses(site_id, creds)
- process_site(site, chunks, creds, filters, on_log, on_chunk)

Провайдер выбирается по ключу "provider" в projects.json (дефолт — callibri).
"""

from providers import callibri, calltouch

_REGISTRY = {
    callibri.NAME: callibri,
    calltouch.NAME: calltouch,
}


def get_provider(name=None):
    """Вернуть модуль провайдера по имени. Если name=None — callibri."""
    key = name or "callibri"
    if key not in _REGISTRY:
        raise ValueError(
            f"Неизвестный провайдер: {key}. Доступны: {list(_REGISTRY.keys())}"
        )
    return _REGISTRY[key]


def provider_names():
    return list(_REGISTRY.keys())


def all_providers():
    return list(_REGISTRY.values())
