"""Общий модуль trip-planner: загрузка trip_data.json, проверки и расчёт итогов.

Используется скриптами build_workbook.py, xlsx_to_data.py и check_consistency.py.
Итоги считаются в Python теми же правилами, что и формулы Excel, чтобы
сборку можно было сверить с пересчитанной книгой.
"""

from __future__ import annotations

import ast
import datetime as dt
import json
import operator
import re
from pathlib import Path

YES = "Да"
NO = "Нет"
SELECT_VALUES = {YES, NO}

SECTIONS = ("flights", "hotels", "transport", "meals", "attractions", "other")
ID_SECTIONS = ("transport", "meals", "attractions")

WEEKDAYS_RU = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]
MONTHS_RU_GEN = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]


# ---------------------------------------------------------------- values

_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}


def _eval_node(node):
    if isinstance(node, ast.Expression):
        return _eval_node(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.BinOp) and type(node.op) in _OPS:
        return _OPS[type(node.op)](_eval_node(node.left), _eval_node(node.right))
    if isinstance(node, ast.UnaryOp) and type(node.op) in _OPS:
        return _OPS[type(node.op)](_eval_node(node.operand))
    raise ValueError("unsupported expression")


def is_formula(value) -> bool:
    return isinstance(value, str) and value.strip().startswith("=")


def num(value, default: float = 0.0) -> float:
    """Число из значения JSON: 1600, "1 600", "1,5" или простая формула "=172952+1620"."""
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text.startswith("="):
        expr = text[1:].replace(" ", "")
        if not re.fullmatch(r"[0-9.+\-*/()]+", expr):
            raise ValueError(f"формула не является простой арифметикой: {text}")
        return _eval_node(ast.parse(expr, mode="eval"))
    text = text.replace(" ", "").replace(" ", "").replace(",", ".")
    return float(text)


def parse_date(value) -> dt.date | None:
    if value in (None, ""):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"не удалось разобрать дату: {value!r}")


def ru_date(day: dt.date, with_weekday: bool = True) -> str:
    text = f"{day.day} {MONTHS_RU_GEN[day.month - 1]}"
    if with_weekday:
        text += f", {WEEKDAYS_RU[day.weekday()]}"
    return text


# ---------------------------------------------------------------- loading

def load(path: str | Path) -> dict:
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    apply_defaults(data)
    return data


def save(data: dict, path: str | Path) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2, default=str)
        fh.write("\n")


def apply_defaults(data: dict) -> dict:
    trip = data.setdefault("trip", {})
    trip.setdefault("base_currency", "RUB")
    group = trip.setdefault("group", {})
    group.setdefault("size", 2)
    group.setdefault("label", f"{group['size']} чел.")
    extra = trip.setdefault("extra", {})
    extra.setdefault("size", 0)
    extra.setdefault("label", f"+{extra['size']}" if extra["size"] else "+1")
    for section in SECTIONS:
        data.setdefault(section, [])
    data.setdefault("rates", [{"currency": trip["base_currency"], "rate": 1, "comment": "Базовая валюта"}])
    data.setdefault("days", [])
    data.setdefault("reserve", {"share": 0.0})
    data.setdefault("budget", {})
    data.setdefault("budget_notes", [])
    gs, es = group["size"], extra["size"]
    for item in data["flights"] + data["attractions"] + data["other"]:
        item.setdefault("qty_group", gs)
        item.setdefault("qty_extra", es)
    for section in SECTIONS:
        for item in data[section]:
            item.setdefault("selected", NO)
    return data


def group_size(data: dict) -> int:
    return int(data["trip"]["group"]["size"])


def extra_size(data: dict) -> int:
    return int(data["trip"]["extra"]["size"])


def rate_value(entry: dict) -> float:
    if entry.get("cbr_rate") not in (None, ""):
        return num(entry["cbr_rate"]) * (1 + num(entry.get("markup"), 0.0))
    return num(entry.get("rate"), 0.0)


def rate_map(data: dict) -> dict[str, float]:
    return {entry["currency"]: rate_value(entry) for entry in data["rates"]}


def selected(item: dict) -> bool:
    return item.get("selected") == YES


# ---------------------------------------------------------------- totals

def row_amounts(section: str, item: dict, data: dict) -> tuple[float, float]:
    """Суммы строки в валюте строки: (группа, +1). Для отелей — цены номеров."""
    gs, es = group_size(data), extra_size(data)
    if section == "flights":
        price = num(item.get("price_pp"))
        return price * num(item.get("qty_group"), gs), price * num(item.get("qty_extra"), es)
    if section == "hotels":
        return num(item.get("price_group")), num(item.get("price_extra"))
    if section == "transport":
        return num(item.get("amount_group")), num(item.get("amount_extra"))
    if section == "meals":
        food = num(item.get("food_pp"))
        return food * gs + num(item.get("drinks_group")), food * es + num(item.get("drinks_extra"))
    if section in ("attractions", "other"):
        price = num(item.get("price_pp" if section == "attractions" else "price"))
        return price * num(item.get("qty_group"), gs), price * num(item.get("qty_extra"), es)
    raise KeyError(section)


def totals(data: dict) -> dict:
    """Итоги по выбранным строкам в базовой валюте, как в листе «Бюджет»."""
    rates = rate_map(data)
    result: dict[str, tuple[float, float]] = {}
    for section in ("flights", "hotels", "transport", "meals", "attractions"):
        g = e = 0.0
        for item in data[section]:
            if not selected(item):
                continue
            rate = rates.get(item.get("currency") or data["trip"]["base_currency"], 0.0)
            ag, ae = row_amounts(section, item, data)
            g += ag * rate
            e += ae * rate
        result[section] = (g, e)
    other_rows = []
    for item in data["other"]:
        rate = rates.get(item.get("currency") or data["trip"]["base_currency"], 0.0)
        ag, ae = row_amounts("other", item, data)
        value = (ag * rate, ae * rate) if selected(item) else (0.0, 0.0)
        other_rows.append((item, value))
    result["other_rows"] = other_rows
    reserve = data.get("reserve") or {}
    share = num(reserve.get("share"), 0.0)
    base_sections = ["hotels", "transport", "meals", "attractions"]
    if reserve.get("include_flights"):
        base_sections.append("flights")
    rg = sum(result[s][0] for s in base_sections) + sum(v[0] for it, v in other_rows if it.get("in_reserve"))
    re_ = sum(result[s][1] for s in base_sections) + sum(v[1] for it, v in other_rows if it.get("in_reserve"))
    result["reserve"] = (share * rg, share * re_)
    sections_total = [result[s] for s in ("flights", "hotels", "transport", "meals", "attractions")]
    result["total"] = (
        sum(v[0] for v in sections_total) + sum(v[0] for _, v in other_rows) + result["reserve"][0],
        sum(v[1] for v in sections_total) + sum(v[1] for _, v in other_rows) + result["reserve"][1],
    )
    return result


def format_rub(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ")


def totals_report(data: dict) -> str:
    t = totals(data)
    base = data["trip"]["base_currency"]
    gl, el = data["trip"]["group"]["label"], data["trip"]["extra"]["label"]
    names = {
        "flights": "Перелёты", "hotels": "Проживание", "transport": "Транспорт",
        "meals": "Питание", "attractions": "Достопримечательности",
    }
    lines = [f"| Категория | {base}: {gl} | {base}: {el} |", "|---|---:|---:|"]
    for key, label in names.items():
        g, e = t[key]
        lines.append(f"| {label} | {format_rub(g)} | {format_rub(e)} |")
    for item, (g, e) in t["other_rows"]:
        lines.append(f"| {item.get('category', '')}: {item.get('item', '')} | {format_rub(g)} | {format_rub(e)} |")
    g, e = t["reserve"]
    lines.append(f"| Резерв | {format_rub(g)} | {format_rub(e)} |")
    g, e = t["total"]
    lines.append(f"| **Итого** | **{format_rub(g)}** | **{format_rub(e)}** |")
    return "\n".join(lines)


# ---------------------------------------------------------------- validation

def validate(data: dict) -> tuple[list[str], list[str]]:
    """Ошибки (книгу собирать нельзя) и предупреждения (собрать можно, но проверить)."""
    errors: list[str] = []
    warnings: list[str] = []
    trip = data["trip"]
    currencies = {entry.get("currency") for entry in data["rates"]}
    if trip["base_currency"] not in currencies:
        errors.append(f"В rates нет базовой валюты {trip['base_currency']}")
    for entry in data["rates"]:
        if rate_value(entry) <= 0:
            errors.append(f"Курс {entry.get('currency')} не задан или равен нулю")
    start = end = None
    try:
        start, end = parse_date(trip.get("start_date")), parse_date(trip.get("end_date"))
    except ValueError as exc:
        errors.append(str(exc))
    if start and end and end < start:
        errors.append("trip.end_date раньше trip.start_date")

    for section in SECTIONS:
        for index, item in enumerate(data[section], 1):
            where = f"{section}[{index}]"
            if item.get("selected") not in SELECT_VALUES:
                errors.append(f"{where}: selected должно быть «Да» или «Нет», сейчас {item.get('selected')!r}")
            cur = item.get("currency") or trip["base_currency"]
            if cur not in currencies:
                errors.append(f"{where}: валюта {cur} отсутствует в rates")
            for key in ("date", "checkin", "checkout"):
                if key in item and item[key] not in (None, ""):
                    try:
                        day = parse_date(item[key])
                    except ValueError as exc:
                        errors.append(f"{where}: {exc}")
                        continue
                    if start and end and not (start <= day <= end):
                        warnings.append(f"{where}: дата {day} вне поездки {start}–{end}")
            try:
                row_amounts(section, item, data)
            except ValueError as exc:
                errors.append(f"{where}: {exc}")
            if section in ID_SECTIONS and not item.get("id"):
                errors.append(f"{where}: нет id (нужен для проверки «выбран ровно один вариант»)")

    for section in ID_SECTIONS:
        chosen: dict[str, int] = {}
        for item in data[section]:
            if item.get("id"):
                chosen.setdefault(item["id"], 0)
                chosen[item["id"]] += selected(item)
        for key, count in chosen.items():
            if count > 1:
                errors.append(f"{section}: для {key} выбрано {count} варианта — должен быть один")

    for index, hotel in enumerate(data["hotels"], 1):
        try:
            cin, cout = parse_date(hotel.get("checkin")), parse_date(hotel.get("checkout"))
        except ValueError:
            continue
        if cin and cout and cout <= cin:
            errors.append(f"hotels[{index}]: выезд не позже заезда")
    warnings.extend(hotel_coverage_warnings(data))
    return errors, warnings


def hotel_nights(data: dict, who: str) -> dict[dt.date, list[str]]:
    """Ночи, покрытые выбранными отелями: who = 'group' или 'extra'."""
    price_key = "price_group" if who == "group" else "price_extra"
    nights: dict[dt.date, list[str]] = {}
    for hotel in data["hotels"]:
        if not selected(hotel) or num(hotel.get(price_key)) <= 0:
            continue
        cin, cout = parse_date(hotel.get("checkin")), parse_date(hotel.get("checkout"))
        if not cin or not cout:
            continue
        day = cin
        while day < cout:
            nights.setdefault(day, []).append(hotel.get("name", "?").strip())
            day += dt.timedelta(days=1)
    return nights


def hotel_coverage_warnings(data: dict) -> list[str]:
    warnings = []
    whos = ["group"] + (["extra"] if extra_size(data) > 0 else [])
    labels = {"group": data["trip"]["group"]["label"], "extra": data["trip"]["extra"]["label"]}
    for who in whos:
        nights = hotel_nights(data, who)
        if not nights:
            continue
        first, last = min(nights), max(nights)
        day = first
        while day <= last:
            hotels = nights.get(day, [])
            if not hotels:
                warnings.append(f"Отели ({labels[who]}): ночь {day} не покрыта выбранным отелем")
            elif len(hotels) > 1:
                warnings.append(f"Отели ({labels[who]}): ночь {day} покрыта дважды: {', '.join(hotels)}")
            day += dt.timedelta(days=1)
    return warnings
