#!/usr/bin/env python3
"""Поиск отелей и актуальных тарифов на Островке (ostrovok.ru). Только чтение.

Команды:
    find    — найти отель или регион по названию: id отеля (otahotel_id) и region_id
    region  — отели региона на даты с фильтрами (рейтинг, звёзды, завтрак, расстояние до точки)
    prices  — тарифы конкретных отелей для разного числа взрослых (номер группы и номер «+1»)

Примеры:
    python3 scripts/ostrovok.py find "Sotetsu Fresa Inn Osaka-Namba"
    python3 scripts/ostrovok.py region --region-id 2697 --checkin 2027-04-10 --checkout 2027-04-14 \
        --adults 2 --near 34.6664,135.5009 --max-km 1 --min-rating 8.5 --pages 3
    python3 scripts/ostrovok.py prices --checkin 2027-04-10 --checkout 2027-04-14 \
        --hotel sotetsu_fresa_inn_osakanamba --occupancy 2 --occupancy 1

Цены — за весь срок проживания в RUB. Налоги, не включённые в цену (city_tax и т. п.),
выводятся отдельно: их платят на месте. Островок может вернуть только самый дешёвый
тариф, поэтому завтрак проверяется по полям meal/has_breakfast. Перед бронированием
тариф нужно проверить на сайте глазами.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import time
import urllib.parse
import urllib.request

BASE = "https://ostrovok.ru"
HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json",
    "Content-Type": "application/json",
    "Origin": BASE,
}


def request_json(url: str, payload: dict | None = None) -> dict:
    data = None if payload is None else json.dumps(payload, ensure_ascii=False).encode()
    request = urllib.request.Request(url, data=data, method="POST" if data else "GET", headers=HEADERS)
    with urllib.request.urlopen(request, timeout=45) as response:
        return json.loads(response.read())


def multicomplete(query: str) -> dict:
    q = urllib.parse.quote(query)
    return request_json(f"{BASE}/api/site/multicomplete.json?query={q}&locale=ru")


def find_region_id(hotel_id: str) -> int | None:
    result = multicomplete(hotel_id.replace("_", " "))
    for hotel in result.get("hotels") or []:
        if hotel.get("otahotel_id") == hotel_id:
            return hotel.get("region_id")
    return None


def serp(region_id: int, checkin: str, checkout: str, adults: int, hotel_id: str | None = None, page: int = 1) -> dict:
    params = {
        "arrival_date": checkin,
        "departure_date": checkout,
        "currency": "RUB",
        "default_sort": "popularity",
        "language": "ru",
        "paxes": [{"adults": adults}],
        "region_id": region_id,
        "member_deals_enabled": False,
        "guru_auth_prices_enabled": False,
    }
    if hotel_id:
        params["hotel"] = hotel_id
        params["selected_hotel"] = hotel_id
    body = urllib.parse.quote(json.dumps({"session_params": params}, ensure_ascii=False, separators=(",", ":")))
    first = request_json(f"{BASE}/hotel/search/v2/site/serp?body={body}")
    session_id = first["session_info"]["session"]["id"]
    url = f"{BASE}/hotel/search/v2/site/serp?session={urllib.parse.quote(session_id)}"
    result = first
    for _ in range(15):
        payload = {"session_params": params, "page": page, "filters": None, "sort": None,
                   "session_id": session_id, "map_hotels": False}
        result = request_json(url, payload)
        if result.get("search_finished"):
            break
        time.sleep(1)
    return result


def parse_rate(rate: dict) -> dict:
    payment = (rate.get("payment_options", {}).get("payment_types") or [{}])[0]
    room = (rate.get("rooms") or [{}])[0]
    meals = room.get("meal_data", {}).get("meals") or []
    taxes = payment.get("tax_data", {}).get("taxes") or []
    amount = payment.get("show_amount") or payment.get("amount")
    return {
        "amount_rub": float(amount) if amount not in (None, "") else None,
        "room": room.get("room_name"),
        "meal": room.get("meal"),
        "has_breakfast": any(item.get("has_breakfast") for item in meals),
        "free_cancellation_before": rate.get("cancellation_info", {}).get("free_cancellation_before"),
        "taxes_not_included": [
            {"name": t.get("trans_key"), "amount": t.get("amount"), "currency": t.get("currency_code")}
            for t in taxes if not t.get("included_in_show")
        ],
    }


def distance_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


# ---------------------------------------------------------------- commands

def cmd_find(args) -> list[dict]:
    result = multicomplete(args.query)
    out = []
    for region in (result.get("regions") or [])[: args.limit]:
        out.append({"type": "region", "region_id": region.get("id"), "name": region.get("name"),
                    "country": region.get("country_name") or region.get("country_code")})
    for hotel in (result.get("hotels") or [])[: args.limit]:
        out.append({"type": "hotel", "hotel_id": hotel.get("otahotel_id"), "name": hotel.get("hotel_name"),
                    "region_id": hotel.get("region_id"), "region": hotel.get("region_name"),
                    "address": hotel.get("hotel_address")})
    return out


def cmd_region(args) -> list[dict]:
    near = tuple(float(x) for x in args.near.split(",")) if args.near else None
    rows = []
    for page in range(1, args.pages + 1):
        result = serp(args.region_id, args.checkin, args.checkout, args.adults, page=page)
        for hotel in result.get("hotels") or []:
            sv = hotel.get("static_vm") or {}
            rates = [parse_rate(r) for r in hotel.get("rates") or []]
            rates = [r for r in rates if r["amount_rub"] is not None]
            if not rates:
                continue
            best = min(rates, key=lambda r: r["amount_rub"])
            with_breakfast = [r for r in rates if r["has_breakfast"]]
            rating = (sv.get("rating") or {}).get("total")
            stars = (sv.get("star_rating") or 0) / 10
            dist = None
            if near and sv.get("latitude") is not None:
                dist = round(distance_km(near[0], near[1], sv["latitude"], sv["longitude"]), 2)
            row = {
                "hotel_id": hotel.get("ota_hotel_id"),
                "name": sv.get("name_en") or sv.get("name"),
                "stars": stars,
                "rating": rating,
                "reviews": (sv.get("rating") or {}).get("count"),
                "address": sv.get("address"),
                "lat": sv.get("latitude"),
                "lon": sv.get("longitude"),
                "distance_km": dist,
                "min_price_rub": best["amount_rub"],
                "min_price_breakfast_rub": min(r["amount_rub"] for r in with_breakfast) if with_breakfast else None,
                "free_cancellation_before": best["free_cancellation_before"],
                "taxes_not_included": best["taxes_not_included"],
            }
            if args.min_rating and (rating or 0) < args.min_rating:
                continue
            if args.min_stars and stars < args.min_stars:
                continue
            if args.breakfast and row["min_price_breakfast_rub"] is None:
                continue
            if args.max_price and row["min_price_rub"] > args.max_price:
                continue
            if near and args.max_km and (dist is None or dist > args.max_km):
                continue
            rows.append(row)
    key = (lambda r: r["distance_km"] or 0) if args.sort == "distance" else (lambda r: r["min_price_rub"])
    if args.sort == "rating":
        key = lambda r: -(r["rating"] or 0)  # noqa: E731
    rows.sort(key=key)
    return rows[: args.limit]


def cmd_prices(args) -> list[dict]:
    output = []
    for hotel_spec in args.hotel:
        hotel_id, _, region = hotel_spec.partition("@")
        region_id = int(region) if region else (args.region_id or find_region_id(hotel_id))
        if not region_id:
            output.append({"hotel_id": hotel_id, "error": "region_id не найден: укажите --region-id или hotel@region"})
            continue
        for adults in args.occupancy or [2]:
            try:
                result = serp(region_id, args.checkin, args.checkout, adults, hotel_id=hotel_id)
                hotel = next((h for h in result.get("hotels") or [] if h.get("ota_hotel_id") == hotel_id), None)
                if hotel is None:
                    output.append({"hotel_id": hotel_id, "adults": adults, "error": "нет свободных номеров или отель не найден"})
                    continue
                output.append({
                    "hotel_id": hotel_id,
                    "hotel_name": (hotel.get("static_vm") or {}).get("name"),
                    "region_id": region_id,
                    "adults": adults,
                    "checkin": args.checkin,
                    "checkout": args.checkout,
                    "rates": [parse_rate(r) for r in hotel.get("rates") or []],
                    "checked_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                })
            except Exception as exc:  # остальные проверки продолжаются
                output.append({"hotel_id": hotel_id, "adults": adults, "error": f"{type(exc).__name__}: {exc}"})
    return output


def print_table(rows: list[dict]) -> None:
    if not rows:
        print("Ничего не найдено.")
        return
    print("| # | Отель | id | ★ | Рейтинг | км | Мин. цена, RUB | С завтраком, RUB | Бесплатная отмена до |")
    print("|---|---|---|---|---|---|---:|---:|---|")
    for i, r in enumerate(rows, 1):
        dist = "" if r["distance_km"] is None else f"{r['distance_km']:.2f}"
        bf = "" if r["min_price_breakfast_rub"] is None else f"{r['min_price_breakfast_rub']:,.0f}"
        print(f"| {i} | {r['name']} | `{r['hotel_id']}` | {r['stars']:g} | {r['rating'] or ''} ({r['reviews'] or 0}) "
              f"| {dist} | {r['min_price_rub']:,.0f} | {bf} | {r['free_cancellation_before'] or ''} |")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("find")
    p.add_argument("query")
    p.add_argument("--limit", type=int, default=8)

    p = sub.add_parser("region")
    p.add_argument("--region-id", type=int, required=True)
    p.add_argument("--checkin", required=True)
    p.add_argument("--checkout", required=True)
    p.add_argument("--adults", type=int, default=2)
    p.add_argument("--pages", type=int, default=2, help="страниц по 20 отелей")
    p.add_argument("--near", help="lat,lon точки (станция, центр района)")
    p.add_argument("--max-km", type=float)
    p.add_argument("--min-rating", type=float)
    p.add_argument("--min-stars", type=float)
    p.add_argument("--max-price", type=float, help="максимум за весь срок, RUB")
    p.add_argument("--breakfast", action="store_true", help="только с тарифом с завтраком")
    p.add_argument("--sort", choices=["price", "distance", "rating"], default="price")
    p.add_argument("--limit", type=int, default=20)
    p.add_argument("--json", action="store_true")

    p = sub.add_parser("prices")
    p.add_argument("--checkin", required=True)
    p.add_argument("--checkout", required=True)
    p.add_argument("--hotel", action="append", required=True, help="otahotel_id или otahotel_id@region_id")
    p.add_argument("--occupancy", action="append", type=int, help="взрослых в номере; можно несколько")
    p.add_argument("--region-id", type=int)

    args = parser.parse_args()
    if args.command == "find":
        json.dump(cmd_find(args), sys.stdout, ensure_ascii=False, indent=2)
        print()
    elif args.command == "region":
        rows = cmd_region(args)
        if args.json:
            json.dump(rows, sys.stdout, ensure_ascii=False, indent=2)
            print()
        else:
            print_table(rows)
    else:
        json.dump(cmd_prices(args), sys.stdout, ensure_ascii=False, indent=2)
        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
