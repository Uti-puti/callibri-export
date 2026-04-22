"""
Диагностика подключения к Calltouch: пробует несколько вариантов эндпоинта
списка сайтов и показывает сырой ответ каждого.
"""

import os
import sys
import requests
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding="utf-8")
load_dotenv()

TOKEN = os.getenv("CALLTOUCH_API_ID", "").strip()
if not TOKEN:
    print("ОШИБКА: CALLTOUCH_API_ID не задан в .env")
    sys.exit(1)

print(f"Токен: {TOKEN[:8]}...{TOKEN[-4:]}  (длина {len(TOKEN)})\n")

BASE = "https://api.calltouch.ru"

# Перебираем варианты, которые публично документированы Calltouch
variants = [
    # (описание, метод, URL, params-или-json)
    ("v1: /calls-service/RestAPI/sites", "GET",
     f"{BASE}/calls-service/RestAPI/sites", {"clientApiId": TOKEN}),
    ("v1: /calls-service/RestAPI/account/sites", "GET",
     f"{BASE}/calls-service/RestAPI/account/sites", {"clientApiId": TOKEN}),
    ("v2: /sites-service/sites", "GET",
     f"{BASE}/sites-service/sites", {"clientApiId": TOKEN}),
    ("account-api siteblocks", "GET",
     f"{BASE}/account-api/v1/siteblocks", {"clientApiId": TOKEN}),
    ("account-api sites", "GET",
     f"{BASE}/account-api/v1/sites", {"clientApiId": TOKEN}),
    ("Bearer header, /sites-service/sites", "GET",
     f"{BASE}/sites-service/sites", None),
]

for label, method, url, params in variants:
    print("=" * 70)
    print(f"{label}")
    print(f"{method} {url}")
    try:
        headers = {"Accept": "application/json"}
        if params is None:
            headers["Authorization"] = f"Bearer {TOKEN}"
            resp = requests.get(url, headers=headers, timeout=20)
        else:
            resp = requests.get(url, params=params, headers=headers, timeout=20)
        print(f"Status: {resp.status_code}")
        print(f"Content-Type: {resp.headers.get('Content-Type', '-')}")
        body = resp.text
        if len(body) > 500:
            body = body[:500] + "...(обрезано)"
        print(f"Body: {body}")
    except requests.RequestException as e:
        print(f"Сетевая ошибка: {e}")
    print()
