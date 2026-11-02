#!/usr/bin/env python3
"""Курс валюты ЦБ РФ на дату и готовая запись для trip_data.json.

Пример:
    python3 scripts/cbr_rate.py --currency JPY --markup 0.10
    python3 scripts/cbr_rate.py --currency EUR --currency THB --date 2026-08-07

ЦБ публикует курс за номинал (например, 100 JPY); скрипт переводит его в курс
за 1 единицу. Если на дату курса нет (выходной), ЦБ возвращает последний
опубликованный — фактическая дата берётся из ответа. Только стандартная библиотека.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import urllib.request
import xml.etree.ElementTree as ET

URL = "https://www.cbr.ru/scripts/XML_daily.asp?date_req={date}"


def fetch(date: dt.date | None) -> ET.Element:
    stamp = (date or dt.date.today()).strftime("%d/%m/%Y")
    request = urllib.request.Request(URL.format(date=stamp), headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=30) as response:
        return ET.fromstring(response.read())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--currency", action="append", required=True, help="код валюты, например JPY")
    parser.add_argument("--date", help="YYYY-MM-DD; по умолчанию сегодня")
    parser.add_argument("--markup", type=float, default=0.10, help="надбавка на конвертацию, по умолчанию 0.10")
    args = parser.parse_args()

    date = dt.date.fromisoformat(args.date) if args.date else None
    root = fetch(date)
    actual = dt.datetime.strptime(root.attrib["Date"], "%d.%m.%Y").date()
    by_code = {v.findtext("CharCode"): v for v in root.findall("Valute")}
    entries, missing = [], []
    for code in (c.upper() for c in args.currency):
        valute = by_code.get(code)
        if valute is None:
            missing.append(code)
            continue
        nominal = int(valute.findtext("Nominal"))
        value_text = valute.findtext("Value")
        per_unit = float(value_text.replace(",", ".")) / nominal
        per_unit_text = f"{per_unit:.6f}".rstrip("0").rstrip(".").replace(".", ",")
        comment = f"ЦБ РФ на {actual:%d.%m.%Y}: {nominal} {code} = {value_text} RUB"
        comment += f"; 1 {code} = {per_unit_text} RUB." if nominal != 1 else "."
        entries.append({
            "currency": code,
            "cbr_rate": round(per_unit, 6),
            "markup": args.markup,
            "rate_date": actual.isoformat(),
            "source": "ЦБ РФ",
            "comment": comment + (f" Надбавка {args.markup:.0%} на конвертацию." if args.markup else ""),
        })
        print(f"{comment}  С надбавкой {args.markup:.0%}: {per_unit * (1 + args.markup):.6f} RUB", file=sys.stderr)
    if missing:
        print(f"ЦБ не публикует курс: {', '.join(missing)} — нужен кросс-курс через USD/EUR", file=sys.stderr)
    json.dump(entries, sys.stdout, ensure_ascii=False, indent=2)
    print()
    return 1 if missing else 0


if __name__ == "__main__":
    raise SystemExit(main())
