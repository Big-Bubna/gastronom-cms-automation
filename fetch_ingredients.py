#!/usr/bin/env python3
"""
Скачивает список продуктов с gastronom.ru/product/letter/...
и сохраняет в ingredients.json.

Не требует авторизации — только публичные страницы.

Запуск:
    python fetch_ingredients.py
"""

import re
import sys
import time
import json
from pathlib import Path
from urllib.request import urlopen, Request

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE     = "https://www.gastronom.ru"
OUT_FILE = Path("ingredients.json")

LETTERS = {
    "А": "a",  "Б": "b",  "В": "v",   "Г": "g",  "Д": "d",
    "Е": "e",  "Ё": "yo", "Ж": "zh",  "З": "z",  "И": "i",
    "Й": "j",  "К": "k",  "Л": "l",   "М": "m",  "Н": "n",
    "О": "o",  "П": "p",  "Р": "r",   "С": "s",  "Т": "t",
    "У": "u",  "Ф": "f",  "Х": "kh",  "Ц": "cz", "Ч": "ch",
    "Ш": "sh", "Щ": "shh","Э": "eh",  "Ю": "yu", "Я": "ya",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


def fetch_html(url: str) -> str:
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=15) as resp:
        return resp.read().decode("utf-8", errors="replace")


def get_max_page(html: str) -> int:
    nums = re.findall(r'[?&]page=(\d+)', html)
    return max((int(n) for n in nums), default=1)


def parse_products(html: str) -> list:
    # Ссылки вида href="/product/slug" с текстовым содержимым (не img)
    pattern = r'<a[^>]+href="/product/[^"?#]+"[^>]*>\s*([^<]{2,}?)\s*</a>'
    names = []
    for m in re.finditer(pattern, html, re.DOTALL):
        name = m.group(1).strip()
        if name:
            names.append(name)
    return names


def scrape_letter(letter: str, slug: str, all_names: set) -> int:
    url1 = f"{BASE}/product/letter/{slug}"
    try:
        html = fetch_html(url1)
    except Exception as e:
        print(f"  Ошибка на {letter}: {e}")
        return 0

    all_names.update(parse_products(html))
    max_page = get_max_page(html)

    for pg in range(2, max_page + 1):
        try:
            h = fetch_html(f"{BASE}/product/letter/{slug}?page={pg}")
            all_names.update(parse_products(h))
        except Exception as e:
            print(f"    стр.{pg} ошибка: {e}")
        time.sleep(0.2)

    return max_page


def main():
    print("Скачиваю продукты с gastronom.ru/product/letter/...\n")
    all_names: set[str] = set()

    for letter, slug in LETTERS.items():
        n_before = len(all_names)
        pages = scrape_letter(letter, slug, all_names)
        added = len(all_names) - n_before
        print(f"  {letter}: {pages} стр., +{added}  (итого {len(all_names)})")
        time.sleep(0.3)

    names_sorted = sorted(all_names, key=str.lower)
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(names_sorted, f, ensure_ascii=False, indent=2)

    print(f"\nГотово! Сохранено {len(names_sorted)} продуктов в {OUT_FILE}")


if __name__ == "__main__":
    main()
