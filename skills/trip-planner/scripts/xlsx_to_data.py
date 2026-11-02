#!/usr/bin/env python3
"""Прочитать книгу расчётов поездки и сохранить данные в trip_data.json.

Книга не изменяется. Столбцы ищутся по заголовкам, поэтому читаются и книги,
собранные build_workbook.py, и старый формат «Расчеты_Япония.xlsx»
(без столбцов «Валюта»: валюта берётся из заголовка вроде «JPY: 2 чел.»).

Примеры:
    # перенести ручные правки Excel в данные, сохранив дни и параметры поездки
    bash scripts/run.sh xlsx_to_data.py --xlsx Расчеты_Италия.xlsx --base trip_data.json --out trip_data.json

    # импорт старой книги без trip_data.json
    bash scripts/run.sh xlsx_to_data.py --xlsx Расчеты_Япония.xlsx --out trip_data.json \
        --country Япония --start 2026-10-16 --end 2026-10-31 --group-size 2 --extra-size 1 --local-currency JPY
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tripdata as td  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

SKIP = "__skip__"
CODE_PREFIX = re.compile(r"^([a-z]{3}):")
CODE_SUFFIX = re.compile(r",\s*([a-z]{3})$")


def norm(text) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip().lower().replace("ё", "е"))


# Правила: (поле или [поле группы, поле +1], проверка заголовка). Первое совпадение выигрывает.
def rules_for(sheet: str, base: str):
    b = base.lower()
    rub = lambda h: h.startswith(f"{b}:")  # noqa: E731 — расчётные столбцы в базовой валюте
    code = lambda h: bool(CODE_PREFIX.match(h)) and not rub(h)  # noqa: E731
    common_tail = [("selected", lambda h: h == "выбран"), (SKIP, lambda h: h.startswith("проверка"))]
    if sheet == "Курсы":
        return [
            ("currency", lambda h: h == "валюта"),
            ("rate", lambda h: h.startswith("курс к")),
            ("cbr_rate", lambda h: h.startswith("курс источника")),
            ("markup", lambda h: h.startswith("надбавка")),
            ("rate_date", lambda h: h.startswith("дата курса")),
            ("comment", lambda h: h.startswith("комментарий") or h.startswith("источник")),
        ]
    if sheet == "Перелеты":
        return [
            ("option", lambda h: h == "вариант"),
            ("details", lambda h: h.startswith("время")),
            ("baggage", lambda h: h.startswith("багаж")),
            ("currency", lambda h: h == "валюта"),
            ("price_pp", lambda h: h.startswith("цена 1 чел")),
            (["qty_group", "qty_extra"], lambda h: h.startswith("количество")),
            (SKIP, lambda h: rub(h) or h.startswith("сумма")),
            ("status", lambda h: h == "статус"),
            ("comment", lambda h: h.startswith("источник") or h.startswith("комментарий")),
        ] + common_tail
    if sheet == "Отели":
        return [
            ("city", lambda h: h.startswith("город")),
            ("checkin", lambda h: h == "заезд"),
            ("checkout", lambda h: h == "выезд"),
            (SKIP, lambda h: h == "ночей"),
            ("name", lambda h: h == "отель"),
            ("area", lambda h: h.startswith("район")),
            ("stars", lambda h: h.startswith("звезд")),
            ("currency", lambda h: h == "валюта"),
            (["price_group", "price_extra"], lambda h: h.startswith("цена: номер") or h.startswith(f"{b}: номер")),
            (SKIP, lambda h: rub(h)),
            (["booking_group", "booking_extra"], lambda h: h.startswith("бронь") or h.startswith("островок")),
            ("status", lambda h: h == "статус"),
            ("comment", lambda h: h.startswith("тариф")),
            ("address_local", lambda h: h.startswith("адрес на")),
            ("address", lambda h: h == "адрес"),
            ("phone", lambda h: h.startswith("телефон")),
        ] + common_tail
    if sheet == "Транспорт":
        return [
            ("id", lambda h: h.startswith("id")),
            ("date", lambda h: h == "дата"),
            ("segment", lambda h: h.startswith("город")),
            ("type", lambda h: h.startswith("тип")),
            ("route", lambda h: h == "маршрут"),
            ("option", lambda h: h == "вариант"),
            ("fare", lambda h: h.startswith("тариф")),
            ("currency", lambda h: h == "валюта"),
            (["calc_group", "calc_extra"], lambda h: h.startswith("расчет")),
            (SKIP, lambda h: rub(h)),
            (["amount_group", "amount_extra"], lambda h: h.startswith("сумма:") or code(h)),
            ("status", lambda h: h == "статус"),
            ("comment", lambda h: h.startswith("комментарий")),
        ] + common_tail
    if sheet == "Питание":
        return [
            ("id", lambda h: h.startswith("id")),
            ("date", lambda h: h == "дата"),
            ("area", lambda h: h.startswith("город")),
            ("meal", lambda h: h.startswith("прием пищи")),
            ("venue", lambda h: h.startswith("заведение")),
            ("level", lambda h: h.startswith("уровень")),
            ("why", lambda h: h.startswith("почему")),
            ("currency", lambda h: h == "валюта"),
            ("food_pp", lambda h: h.startswith("еда на 1 чел")),
            (["drinks_group", "drinks_extra"], lambda h: h.startswith("алкоголь")),
            (SKIP, lambda h: rub(h) or h.startswith("сумма:") or code(h)),
            ("status", lambda h: h == "статус"),
            ("source", lambda h: h.startswith("цена")),
        ] + common_tail
    if sheet == "Достопримечательности":
        return [
            ("id", lambda h: h == "id"),
            ("date", lambda h: h == "дата"),
            ("city", lambda h: h.startswith("город")),
            ("name", lambda h: h.startswith("достопримечательность")),
            ("ticket", lambda h: h.startswith("вариант")),
            ("currency", lambda h: h == "валюта"),
            ("price_pp", lambda h: h.startswith("цена 1 чел")),
            (["qty_group", "qty_extra"], lambda h: h.startswith("количество")),
            (SKIP, lambda h: rub(h) or h.startswith("сумма:") or code(h)),
            ("price_status", lambda h: h.startswith("статус цены")),
            ("status", lambda h: h == "статус"),
            ("comment", lambda h: h.startswith("источник") or h.startswith("комментарий")),
        ] + common_tail
    if sheet == "Прочее":
        return [
            ("category", lambda h: h.startswith("категория")),
            ("item", lambda h: h.startswith("позиция")),
            ("currency", lambda h: h == "валюта"),
            ("price", lambda h: h.startswith("цена")),
            (["qty_group", "qty_extra"], lambda h: h.startswith("количество")),
            (SKIP, lambda h: rub(h) or h.startswith("сумма")),
            ("in_reserve", lambda h: h.startswith("в резерв")),
            ("comment", lambda h: h.startswith("комментарий")),
        ] + common_tail
    raise KeyError(sheet)


def map_columns(ws, sheet: str, base: str) -> tuple[dict[str, int], dict[str, str]]:
    """Поле → номер столбца; поле → валюта из заголовка (старый формат)."""
    fields: dict[str, int] = {}
    header_currency: dict[str, str] = {}
    used_pairs: dict[int, int] = {}
    rules = rules_for(sheet, base)
    for col in range(1, ws.max_column + 1):
        h = norm(ws.cell(row=1, column=col).value)
        if not h:
            continue
        for index, (field, test) in enumerate(rules):
            if not test(h):
                continue
            if field == SKIP:
                break
            if isinstance(field, list):
                pos = used_pairs.get(index, 0)
                if pos >= len(field):
                    break
                used_pairs[index] = pos + 1
                field = field[pos]
            fields.setdefault(field, col)
            m = CODE_PREFIX.match(h) or CODE_SUFFIX.search(h)
            if m:
                header_currency[field] = m.group(1).upper()
            break
    return fields, header_currency


def clean(value):
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, str):
        value = value.strip()
        return value or None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def input_value(cell_f, cell_v):
    """Значение входной ячейки: простая формула сохраняется как строка, иначе кэш."""
    raw = cell_f.value
    if td.is_formula(raw):
        try:
            td.num(raw)
            return raw.strip()
        except ValueError:
            return clean(cell_v.value)
    return clean(raw)


def read_rows(wb_f, wb_v, sheet: str, base: str, default_currency: str):
    if sheet not in wb_f.sheetnames:
        return [], {}
    ws_f, ws_v = wb_f[sheet], wb_v[sheet]
    fields, header_currency = map_columns(ws_f, sheet, base)
    rows = []
    for r in range(2, ws_f.max_row + 1):
        first = ws_f.cell(row=r, column=1).value
        if first is None and all(ws_f.cell(row=r, column=c).value is None for c in fields.values()):
            continue
        if isinstance(first, str) and norm(first).startswith("итого"):
            continue
        item = {"_row": r}
        for field, col in fields.items():
            value = input_value(ws_f.cell(row=r, column=col), ws_v.cell(row=r, column=col))
            if value is not None:
                item[field] = value
            link = ws_f.cell(row=r, column=col).hyperlink
            if link is not None and link.target and field.startswith("booking"):
                item[field.replace("booking", "link")] = link.target
        if "currency" not in item and sheet != "Курсы":
            cur = next((header_currency[k] for k in header_currency), None)
            item["currency"] = cur or default_currency
        if sheet != "Курсы":
            sel = norm(item.get("selected"))
            item["selected"] = td.YES if (sel == "да" or ("selected" not in fields)) else td.NO
        rows.append(item)
    return rows, fields


def parse_rates(rows: list[dict]) -> list[dict]:
    rates = []
    for row in rows:
        entry = {"currency": row["currency"]}
        raw = row.get("rate")
        if row.get("cbr_rate") is not None:
            entry["cbr_rate"] = row["cbr_rate"]
            entry["markup"] = row.get("markup", 0)
        elif isinstance(raw, str):
            text = raw.replace(" ", "")
            factor = re.fullmatch(r"=\(?([\d.]+)\)?\*([\d.]+)", text)  # =(0.516121)*1.1
            plus = re.fullmatch(r"=\(?([\d.]+)\)?\*\(1\+([\d.]+)\)", text)  # =(0.516121)*(1+0.1)
            if plus:
                entry["cbr_rate"], entry["markup"] = float(plus.group(1)), float(plus.group(2))
            elif factor and float(factor.group(2)) >= 1:
                entry["cbr_rate"] = float(factor.group(1))
                entry["markup"] = round(float(factor.group(2)) - 1, 6)
            else:
                entry["rate"] = raw
        elif raw is not None:
            entry["rate"] = raw
        else:
            continue  # валюта без курса (например, EUR «заполнить при необходимости»)
        if row.get("rate_date"):
            entry["rate_date"] = row["rate_date"]
        comment = row.get("comment")
        if comment:
            entry["comment"] = comment
            m = re.search(r"на (\d{2}\.\d{2}\.\d{4})", comment)
            if m and "rate_date" not in entry:
                entry["rate_date"] = td.parse_date(m.group(1)).isoformat()
        rates.append(entry)
    return rates


def strip_rows(items: list[dict]) -> list[dict]:
    return [{k: v for k, v in item.items() if not k.startswith("_")} for item in items]


def read_budget(wb_f, other_rows: list[dict], reserve_row: int | None, data: dict) -> None:
    if "Бюджет" not in wb_f.sheetnames:
        return
    ws = wb_f["Бюджет"]
    section_by_sheet = {
        "перелеты": "flights", "отели": "hotels", "транспорт": "transport",
        "питание": "meals", "достопримечательности": "attractions",
    }
    by_row = {item["_row"]: item for item in other_rows}
    for r in range(2, ws.max_row + 1):
        category = clean(ws.cell(row=r, column=1).value)
        if not category or norm(category).startswith("итого"):
            continue
        meta = {
            "category": category,
            "item": clean(ws.cell(row=r, column=2).value),
            "stage": clean(ws.cell(row=r, column=3).value),
            "status": clean(ws.cell(row=r, column=7).value),
            "comment": clean(ws.cell(row=r, column=8).value),
        }
        meta = {k: v for k, v in meta.items() if v is not None}
        formula = ws.cell(row=r, column=5).value
        m = re.match(r"^=\s*'?([^'!]+)'?!\$?[A-Z]+\$?(\d+)$", str(formula or ""))
        if not m:
            note = dict(meta)
            for column, key in ((5, "amount_group"), (6, "amount_extra")):
                value = ws.cell(row=r, column=column).value
                if isinstance(value, (int, float)):
                    note[key] = value
            data["budget_notes"].append(note)
            continue
        sheet, row = norm(m.group(1)), int(m.group(2))
        if sheet in section_by_sheet:
            data["budget"][section_by_sheet[sheet]] = meta
        elif sheet == "прочее" and row == reserve_row:
            data["reserve"]["budget_status"] = meta.get("status")
            data["reserve"]["budget_comment"] = meta.get("comment")
        elif sheet == "прочее" and row in by_row:
            item = by_row[row]
            item["budget_stage"] = meta.get("stage")
            item["budget_status"] = meta.get("status")
            item["budget_comment"] = meta.get("comment")


def extract(xlsx: Path, trip: dict) -> dict:
    wb_f = load_workbook(xlsx, data_only=False)
    wb_v = load_workbook(xlsx, data_only=True)
    base = trip.get("base_currency", "RUB")
    local = trip.get("local_currency") or base
    data = {"schema_version": 1, "trip": trip, "budget": {}, "budget_notes": [], "reserve": {}}

    rate_rows, _ = read_rows(wb_f, wb_v, "Курсы", base, base)
    data["rates"] = parse_rates(rate_rows)

    flights, _ = read_rows(wb_f, wb_v, "Перелеты", base, base)
    data["flights"] = strip_rows(flights)
    for name, section in (("Отели", "hotels"), ("Транспорт", "transport"), ("Питание", "meals"),
                          ("Достопримечательности", "attractions")):
        rows, _ = read_rows(wb_f, wb_v, name, base, base if section == "hotels" else local)
        data[section] = strip_rows(rows)

    other, fields = read_rows(wb_f, wb_v, "Прочее", base, local)
    reserve_row = None
    kept = []
    for item in other:
        if norm(item.get("category")) == "резерв":
            reserve_row = item["_row"]
            data["reserve"]["share"] = td.num(item.get("price"), 0)
            data["reserve"]["comment"] = item.get("comment")
            ws = wb_f["Прочее"]
            formula = ""
            for col in range(1, ws.max_column + 1):
                value = ws.cell(row=reserve_row, column=col).value
                if td.is_formula(value) and "!" in value:
                    formula = value
                    break
            data["reserve"]["include_flights"] = "Перелеты!" in formula
            local_refs = {int(x) for x in re.findall(r"(?<![!$A-Z])\$?[A-Z]{1,2}\$?(\d+)", formula)}
            for other_item in other:
                if other_item["_row"] in local_refs and other_item["_row"] != reserve_row:
                    other_item["in_reserve"] = True
            continue
        kept.append(item)
    for item in kept:
        if "in_reserve" in fields:
            item["in_reserve"] = norm(item.get("in_reserve")) == "да" or item.get("in_reserve") is True
        else:
            item.setdefault("in_reserve", False)
    read_budget(wb_f, kept, reserve_row, data)
    data["other"] = strip_rows(kept)
    data["reserve"] = {k: v for k, v in data["reserve"].items() if v is not None}
    data["days"] = []
    return data


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--xlsx", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--base", help="существующий trip_data.json: из него берутся trip, days и прочие поля")
    parser.add_argument("--country")
    parser.add_argument("--title")
    parser.add_argument("--start")
    parser.add_argument("--end")
    parser.add_argument("--base-currency", default="RUB")
    parser.add_argument("--local-currency")
    parser.add_argument("--group-size", type=int, default=2)
    parser.add_argument("--extra-size", type=int, default=0)
    parser.add_argument("--group-label")
    parser.add_argument("--extra-label")
    args = parser.parse_args()

    base_data = td.load(args.base) if args.base else None
    if base_data:
        trip = base_data["trip"]
    else:
        trip = {
            "title": args.title or (f"{args.country}" if args.country else "Поездка"),
            "country": args.country or "",
            "file_suffix": args.country or "",
            "start_date": args.start,
            "end_date": args.end,
            "base_currency": args.base_currency,
            "local_currency": args.local_currency or args.base_currency,
            "group": {"size": args.group_size, "label": args.group_label or f"{args.group_size} чел."},
            "extra": {"size": args.extra_size, "label": args.extra_label or (f"+{args.extra_size}" if args.extra_size else "+1")},
        }
    data = extract(Path(args.xlsx), trip)
    if base_data:
        for key, value in base_data.items():
            if key not in data or key == "days":
                data[key] = value
    td.apply_defaults(data)
    td.save(data, args.out)
    counts = ", ".join(f"{s}: {len(data[s])}" for s in td.SECTIONS)
    print(f"Сохранено {args.out} ({counts}).")
    if base_data:
        old, new = td.totals(base_data)["total"], td.totals(data)["total"]
        print(f"Итог до: {td.format_rub(old[0])} / {td.format_rub(old[1])}; после: {td.format_rub(new[0])} / {td.format_rub(new[1])}")
    print(td.totals_report(data))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
