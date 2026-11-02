#!/usr/bin/env python3
"""Оценка пешеходной дистанции дня по пешим маршрутам OpenStreetMap.

Пример:
    python3 scripts/walk_distance.py --city Tokyo \
        "APA Hotel Nihombashi Bakurocho Ekimae" "Kodemmacho Station" / \
        "Ueno Station" "Tokyo National Museum" "Ameyoko" "Okachimachi Station" / \
        "Akihabara Station" "35.6984,139.7731"

Точки идут по порядку; «/» разрывает цепочку там, где едем транспортом
(переход внутри метро не считается). Точку можно задать названием или
координатами «lat,lon». Геокодирование — Nominatim (не чаще 1 запроса в секунду,
результаты кешируются в ~/.cache/trip-planner/geocode.json), маршруты — OSRM
(routing.openstreetmap.de, профиль foot). Итог округляется до 0,5 км и
считается оценкой: для программы его стоит сверить с Google Maps.
Если маршрутизатор недоступен, используется прямая × 1,3.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

UA = "trip-planner-skill/1.0 (personal trip planning)"
CACHE = Path.home() / ".cache" / "trip-planner" / "geocode.json"
COORD = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*$")


def get_json(url: str) -> object:
    request = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read())


def load_cache() -> dict:
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def geocode(name: str, city: str | None, cache: dict) -> tuple[float, float, str]:
    m = COORD.match(name)
    if m:
        return float(m.group(1)), float(m.group(2)), name
    query = f"{name}, {city}" if city else name
    if query in cache:
        lat, lon, label = cache[query]
        return lat, lon, label
    # Остановки и станции OSM находит лучше с явным типом: «Karaköy tram stop, Istanbul».
    variants = [query, name] + [f"{name} {kind}" + (f", {city}" if city else "")
                                for kind in ("station", "tram stop", "metro station", "bus stop", "pier", "ferry terminal")]
    for variant in dict.fromkeys(variants):
        time.sleep(1.1)
        url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
            {"q": variant, "format": "jsonv2", "limit": 1})
        result = get_json(url)
        if result:
            item = result[0]
            value = (float(item["lat"]), float(item["lon"]), item.get("display_name", variant)[:90])
            cache[query] = value
            return value
    raise LookupError(f"не найдено: {query} — уточните название (латиницей, как в OSM) или задайте координаты lat,lon")


def straight_km(a, b) -> float:
    r = 6371.0
    p1, p2 = math.radians(a[0]), math.radians(b[0])
    dp, dl = p2 - p1, math.radians(b[1] - a[1])
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(h))


def route_km(a, b) -> tuple[float, float | None, str]:
    url = (f"https://routing.openstreetmap.de/routed-foot/route/v1/driving/"
           f"{a[1]},{a[0]};{b[1]},{b[0]}?overview=false")
    try:
        data = get_json(url)
        route = data["routes"][0]
        return route["distance"] / 1000, route["duration"] / 60, "OSRM foot"
    except Exception:  # маршрутизатор недоступен — грубая оценка
        return straight_km(a, b) * 1.3, None, "прямая × 1,3"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("points", nargs="+", help="точки по порядку; «/» — разрыв (едем транспортом)")
    parser.add_argument("--city", help="добавляется к названиям при поиске, например Tokyo")
    args = parser.parse_args()

    cache = load_cache()
    chains, chain = [], []
    for token in args.points:
        if token == "/":
            if chain:
                chains.append(chain)
            chain = []
        else:
            chain.append(token)
    if chain:
        chains.append(chain)

    total = 0.0
    try:
        resolved = [[(name, geocode(name, args.city, cache)) for name in chain] for chain in chains]
    except LookupError as exc:
        print(f"ОШИБКА: {exc}. Координаты отеля есть в выводе ostrovok.py region (lat/lon) или в Google Maps.")
        return 2
    finally:
        CACHE.parent.mkdir(parents=True, exist_ok=True)
        CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    print("| Отрезок | км | мин | метод |")
    print("|---|---:|---:|---|")
    for pairs in resolved:
        for (a_name, a), (b_name, b) in zip(pairs, pairs[1:]):
            km, minutes, method = route_km(a, b)
            total += km
            mins = "" if minutes is None else f"{minutes:.0f}"
            print(f"| {a_name} → {b_name} | {km:.2f} | {mins} | {method} |")
    rounded = round(total * 2) / 2
    print(f"\nИтого: {total:.2f} км → в программу: около **{rounded:g} км**".replace(".", ","))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
