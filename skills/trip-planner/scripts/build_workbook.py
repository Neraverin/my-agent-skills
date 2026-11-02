#!/usr/bin/env python3
"""Собрать Excel-бюджет поездки из trip_data.json.

Пример:
    bash scripts/run.sh build_workbook.py --data trip_data.json --out Расчеты_Италия.xlsx

Книга повторяет структуру расчётов поездки в Японию 2026: «Бюджет», «Курсы»,
«Перелеты», «Отели», «Транспорт», «Питание», «Достопримечательности», «Прочее».
Суммы в рублях считаются формулами через курс валюты строки; итоги — только по
строкам с «Выбран» = «Да»; группа и «+1» считаются раздельно.

Защита ручных правок: рядом с книгой хранится .trip_build.json с хешем
собранного файла. Если книга изменилась после последней сборки, скрипт
останавливается (код 3) — дальше изменения выдаются пользователю сводкой ячеек.
--force перезаписывает книгу только по явному решению пользователя.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tripdata as td  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import FormulaRule  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

SIDECAR = ".trip_build.json"
RATES_RANGE = "Курсы!$A$2:$B$40"
MONEY = "#,##0"
DATE = "DD.MM.YYYY"
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="F8CBAD")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
YES_FILL = PatternFill("solid", fgColor="E2EFDA")


@dataclass
class Col:
    key: str
    header: str
    width: float = 14
    fmt: str | None = None
    formula: str | None = None  # шаблон: {r} — строка, {<key>} — буква столбца, {rate} — курс строки
    kind: str = "text"  # text | date | number | select | int
    extra: bool = False  # столбец относится к «+1» и скрывается, если extra.size = 0


def sheet_specs(gl: str, el: str, gs: int, es: int, base: str) -> dict[str, list[Col]]:
    sel_check = (
        '=IF(COUNTIFS(${id}$2:${id}${last},{id}{r},${selected}$2:${selected}${last},"Да")=1,"OK",'
        'IF(COUNTIFS(${id}$2:${id}${last},{id}{r},${selected}$2:${selected}${last},"Да")=0,"Не выбрано","Выбрано >1"))'
    )
    return {
        "Перелеты": [
            Col("option", "Вариант", 16),
            Col("details", "Время / перевозчики", 40),
            Col("baggage", "Багаж", 18),
            Col("currency", "Валюта", 8),
            Col("price_pp", "Цена 1 чел.", 13, MONEY, kind="number"),
            Col("qty_group", f"Количество: {gl}", 11, kind="number"),
            Col("qty_extra", f"Количество: {el}", 11, kind="number", extra=True),
            Col("rub_group", f"{base}: {gl}", 14, MONEY, '=IF({selected}{r}="Да",{price_pp}{r}*{qty_group}{r}*{rate},0)'),
            Col("rub_extra", f"{base}: {el}", 14, MONEY, '=IF({selected}{r}="Да",{price_pp}{r}*{qty_extra}{r}*{rate},0)', extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("status", "Статус", 14),
            Col("comment", "Источник / комментарий", 70),
        ],
        "Отели": [
            Col("city", "Город", 12),
            Col("checkin", "Заезд", 12, DATE, kind="date"),
            Col("checkout", "Выезд", 12, DATE, kind="date"),
            Col("nights", "Ночей", 7, None, "={checkout}{r}-{checkin}{r}"),
            Col("name", "Отель", 34),
            Col("area", "Район / транспорт", 28),
            Col("stars", "Звёзды", 8),
            Col("currency", "Валюта", 8),
            Col("price_group", f"Цена: номер {gl}", 14, MONEY, kind="number"),
            Col("price_extra", f"Цена: номер {el}", 14, MONEY, kind="number", extra=True),
            Col("rub_group", f"{base}: {gl}", 14, MONEY, "={price_group}{r}*{rate}"),
            Col("rub_extra", f"{base}: {el}", 14, MONEY, "={price_extra}{r}*{rate}", extra=True),
            Col("rub_total", f"{base}: итого", 14, MONEY, "={rub_group}{r}+{rub_extra}{r}"),
            Col("booking_group", f"Бронь / тариф: {gl}", 24),
            Col("booking_extra", f"Бронь / тариф: {el}", 24, extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("status", "Статус", 14),
            Col("comment", "Тариф / комментарий", 50),
            Col("address", "Адрес", 40),
            Col("address_local", "Адрес на местном языке", 40),
            Col("phone", "Телефон", 18),
        ],
        "Транспорт": [
            Col("id", "ID участка", 9),
            Col("date", "Дата", 11, DATE, kind="date"),
            Col("segment", "Город / этап", 18),
            Col("type", "Тип транспорта", 16),
            Col("route", "Маршрут", 40),
            Col("option", "Вариант", 26),
            Col("fare", "Тариф / дата", 22),
            Col("currency", "Валюта", 8),
            Col("calc_group", f"Расчёт: {gl}", 14),
            Col("amount_group", f"Сумма: {gl}", 11, MONEY, kind="number"),
            Col("rub_group", f"{base}: {gl} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{amount_group}{r}*{rate},0)'),
            Col("calc_extra", f"Расчёт: {el}", 14, extra=True),
            Col("amount_extra", f"Сумма: {el}", 11, MONEY, kind="number", extra=True),
            Col("rub_extra", f"{base}: {el} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{amount_extra}{r}*{rate},0)', extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("check", "Проверка выбора", 13, None, sel_check),
            Col("status", "Статус", 12),
            Col("comment", "Комментарий / источник", 50),
        ],
        "Питание": [
            Col("id", "ID приёма пищи", 9),
            Col("date", "Дата", 11, DATE, kind="date"),
            Col("area", "Город / участок маршрута", 24),
            Col("meal", "Приём пищи", 12),
            Col("venue", "Заведение", 30),
            Col("level", "Уровень / формат", 22),
            Col("why", "Почему по пути", 40),
            Col("currency", "Валюта", 8),
            Col("food_pp", "Еда на 1 чел.", 11, MONEY, kind="number"),
            Col("drinks_group", f"Алкоголь / напитки: {gl}", 12, MONEY, kind="number"),
            Col("sum_group", f"Сумма: {gl}", 11, MONEY, "={food_pp}{r}*" + str(gs) + "+{drinks_group}{r}"),
            Col("rub_group", f"{base}: {gl} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{sum_group}{r}*{rate},0)'),
            Col("drinks_extra", f"Алкоголь / напитки: {el}", 12, MONEY, kind="number", extra=True),
            Col("sum_extra", f"Сумма: {el}", 11, MONEY, "={food_pp}{r}*" + str(es) + "+{drinks_extra}{r}", extra=True),
            Col("rub_extra", f"{base}: {el} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{sum_extra}{r}*{rate},0)', extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("check", "Проверка выбора", 13, None, sel_check),
            Col("status", "Статус", 12),
            Col("source", "Цена / источник", 50),
        ],
        "Достопримечательности": [
            Col("id", "ID", 7),
            Col("date", "Дата", 11, DATE, kind="date"),
            Col("city", "Город", 12),
            Col("name", "Достопримечательность", 30),
            Col("ticket", "Вариант / билет", 28),
            Col("currency", "Валюта", 8),
            Col("price_pp", "Цена 1 чел.", 11, MONEY, kind="number"),
            Col("qty_group", f"Количество: {gl}", 11, kind="number"),
            Col("sum_group", f"Сумма: {gl}", 11, MONEY, "={price_pp}{r}*{qty_group}{r}"),
            Col("rub_group", f"{base}: {gl} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{sum_group}{r}*{rate},0)'),
            Col("qty_extra", f"Количество: {el}", 11, kind="number", extra=True),
            Col("sum_extra", f"Сумма: {el}", 11, MONEY, "={price_pp}{r}*{qty_extra}{r}", extra=True),
            Col("rub_extra", f"{base}: {el} (выбранное)", 13, MONEY, '=IF({selected}{r}="Да",{sum_extra}{r}*{rate},0)', extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("check", "Проверка выбора", 13, None, sel_check),
            Col("status", "Статус", 12),
            Col("price_status", "Статус цены", 30),
            Col("comment", "Источник / комментарий", 50),
        ],
        "Прочее": [
            Col("category", "Категория", 14),
            Col("item", "Позиция", 34),
            Col("currency", "Валюта", 8),
            Col("price", "Цена / база", 12, MONEY, kind="number"),
            Col("qty_group", f"Количество: {gl}", 11, kind="number"),
            Col("sum_group", f"Сумма: {gl}", 12, MONEY, "={price}{r}*{qty_group}{r}"),
            Col("rub_group", f"{base}: {gl}", 13, MONEY, '=IF({selected}{r}="Да",{sum_group}{r}*{rate},0)'),
            Col("qty_extra", f"Количество: {el}", 11, kind="number", extra=True),
            Col("sum_extra", f"Сумма: {el}", 12, MONEY, "={price}{r}*{qty_extra}{r}", extra=True),
            Col("rub_extra", f"{base}: {el}", 13, MONEY, '=IF({selected}{r}="Да",{sum_extra}{r}*{rate},0)', extra=True),
            Col("selected", "Выбран", 9, kind="select"),
            Col("in_reserve", "В резерв", 9),
            Col("comment", "Комментарий", 60),
        ],
    }


SECTION_OF_SHEET = {
    "Перелеты": "flights",
    "Отели": "hotels",
    "Транспорт": "transport",
    "Питание": "meals",
    "Достопримечательности": "attractions",
    "Прочее": "other",
}


# ---------------------------------------------------------------- helpers

def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def style_header(ws, cols: list[Col], hide_extra: bool) -> None:
    for index, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=index, value=col.header)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        letter = get_column_letter(index)
        ws.column_dimensions[letter].width = col.width
        if col.extra and hide_extra:
            ws.column_dimensions[letter].hidden = True
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def write_value(cell, col: Col, value) -> None:
    if value is None or value == "":
        return
    if col.kind == "date":
        cell.value = td.parse_date(value)
    elif col.kind == "number":
        if td.is_formula(value):
            cell.value = value.strip()
        else:
            cell.value = td.num(value)
            if cell.value == int(cell.value):
                cell.value = int(cell.value)
    elif col.key == "in_reserve":
        cell.value = td.YES if value in (True, td.YES, "true", "yes") else td.NO
    else:
        cell.value = value
    if col.fmt:
        cell.number_format = col.fmt
    if col.kind == "text":
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def fill_rows(ws, cols: list[Col], items: list[dict], base: str, start_row: int = 2) -> int:
    letters = {col.key: get_column_letter(i) for i, col in enumerate(cols, 1)}
    last = start_row + len(items) - 1
    for offset, item in enumerate(items):
        r = start_row + offset
        rate = f"VLOOKUP({letters['currency']}{r},{RATES_RANGE},2,FALSE)" if "currency" in letters else "1"
        for index, col in enumerate(cols, 1):
            cell = ws.cell(row=r, column=index)
            if col.formula:
                cell.value = col.formula.format(r=r, rate=rate, last=last, **letters)
                if col.fmt:
                    cell.number_format = col.fmt
            elif col.key == "currency":
                cell.value = item.get("currency") or base
            elif col.key in ("booking_group", "booking_extra"):
                write_value(cell, col, item.get(col.key))
                link = item.get(col.key.replace("booking", "link"))
                if link:
                    cell.hyperlink = link
                    cell.font = Font(color="0563C1", underline="single")
            else:
                write_value(cell, col, item.get(col.key))
    return last


def add_selection_rules(ws, cols: list[Col], first: int, last: int) -> None:
    letters = {col.key: get_column_letter(i) for i, col in enumerate(cols, 1)}
    sel = letters["selected"]
    dv = DataValidation(type="list", formula1='"Да,Нет"', allow_blank=True)
    dv.add(f"{sel}{first}:{sel}{max(last, first) + 50}")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add(
        f"{sel}{first}:{sel}{max(last, first)}",
        FormulaRule(formula=[f'{sel}{first}="Да"'], fill=YES_FILL),
    )
    if "check" in letters:
        chk = letters["check"]
        ws.conditional_formatting.add(
            f"{chk}{first}:{chk}{max(last, first)}",
            FormulaRule(formula=[f'{chk}{first}="Выбрано >1"'], fill=BAD_FILL),
        )
        ws.conditional_formatting.add(
            f"{chk}{first}:{chk}{max(last, first)}",
            FormulaRule(formula=[f'{chk}{first}="Не выбрано"'], fill=WARN_FILL),
        )


def total_row(ws, cols: list[Col], row: int, label: str, formulas: dict[str, str]) -> dict[str, str]:
    letters = {col.key: get_column_letter(i) for i, col in enumerate(cols, 1)}
    ws.cell(row=row, column=1, value=label).font = Font(bold=True)
    addresses = {}
    for index in range(1, len(cols) + 1):
        ws.cell(row=row, column=index).fill = TOTAL_FILL
    for key, formula in formulas.items():
        cell = ws[f"{letters[key]}{row}"]
        cell.value = formula
        cell.font = Font(bold=True)
        cell.number_format = MONEY
        addresses[key] = f"{letters[key]}{row}"
    return addresses


def single_currency(items: list[dict], base: str) -> str | None:
    found = {item.get("currency") or base for item in items}
    return found.pop() if len(found) == 1 else None


def date_span(items: list[dict]) -> str:
    days = sorted(td.parse_date(i["date"]) for i in items if td.selected(i) and i.get("date"))
    if not days:
        return ""
    return f"{days[0]:%d.%m}–{days[-1]:%d.%m.%Y}"


def hotel_span(hotels: list[dict]) -> str:
    seen, parts = set(), []
    for h in hotels:
        if not td.selected(h):
            continue
        key = (h.get("city"), h.get("checkin"), h.get("checkout"))
        if key in seen:
            continue
        seen.add(key)
        cin, cout = td.parse_date(h["checkin"]), td.parse_date(h["checkout"])
        parts.append(f"{h.get('city', '')} {cin:%d}–{cout:%d.%m}")
    return " + ".join(parts)


# ---------------------------------------------------------------- build

def build(data: dict) -> Workbook:
    trip = data["trip"]
    base = trip["base_currency"]
    gs, es = td.group_size(data), td.extra_size(data)
    gl, el = trip["group"]["label"], trip["extra"]["label"]
    hide_extra = es == 0
    specs = sheet_specs(gl, el, gs, es, base)

    wb = Workbook()
    ws_budget = wb.active
    ws_budget.title = "Бюджет"

    # Курсы
    ws = wb.create_sheet("Курсы")
    rate_cols = [
        Col("currency", "Валюта", 9), Col("rate", f"Курс к {base}", 13), Col("cbr_rate", "Курс источника", 13),
        Col("markup", "Надбавка на конвертацию", 13), Col("rate_date", "Дата курса", 12),
        Col("comment", "Источник / комментарий", 70),
    ]
    style_header(ws, rate_cols, False)
    for r, entry in enumerate(data["rates"], 2):
        ws.cell(row=r, column=1, value=entry["currency"])
        rate_cell = ws.cell(row=r, column=2)
        if entry.get("cbr_rate") not in (None, ""):
            ws.cell(row=r, column=3, value=td.num(entry["cbr_rate"])).number_format = "0.000000"
            markup = ws.cell(row=r, column=4, value=td.num(entry.get("markup"), 0.0))
            markup.number_format = "0%"
            rate_cell.value = f"=C{r}*(1+D{r})"
        else:
            rate_cell.value = entry["rate"] if td.is_formula(entry.get("rate")) else td.num(entry.get("rate"))
        rate_cell.number_format = "0.000000"
        if entry.get("rate_date"):
            ws.cell(row=r, column=5, value=td.parse_date(entry["rate_date"])).number_format = DATE
        comment = "; ".join(x for x in (entry.get("source"), entry.get("comment")) if x)
        ws.cell(row=r, column=6, value=comment or None).alignment = Alignment(wrap_text=True, vertical="top")

    totals_addr: dict[str, dict[str, str]] = {}

    for sheet, cols in specs.items():
        section = SECTION_OF_SHEET[sheet]
        items = data[section]
        ws = wb.create_sheet(sheet)
        style_header(ws, cols, hide_extra)
        last = fill_rows(ws, cols, items, base)
        if items:
            add_selection_rules(ws, cols, 2, last)
        letters = {col.key: get_column_letter(i) for i, col in enumerate(cols, 1)}
        rng = lambda key: f"{letters[key]}2:{letters[key]}{max(last, 2)}"  # noqa: E731
        sumif = lambda key: f'=SUMIF({rng("selected")},"Да",{rng(key)})'  # noqa: E731
        if sheet == "Прочее":
            last = add_reserve_row(ws, cols, data, last, totals_addr)
            totals_addr["other"] = total_row(ws, cols, last + 2, "Итого", {
                "rub_group": f"=SUM({letters['rub_group']}2:{letters['rub_group']}{last})",
                "rub_extra": f"=SUM({letters['rub_extra']}2:{letters['rub_extra']}{last})",
            })
            continue
        formulas = {
            "rub_group": sumif("rub_group") if sheet in ("Отели",) else f"=SUM({rng('rub_group')})",
            "rub_extra": sumif("rub_extra") if sheet in ("Отели",) else f"=SUM({rng('rub_extra')})",
        }
        if sheet == "Отели":
            formulas["rub_total"] = sumif("rub_total")
        cur = single_currency(items, base)
        if cur and sheet in ("Транспорт",):
            formulas["amount_group"] = sumif("amount_group")
            formulas["amount_extra"] = sumif("amount_extra")
        if cur and sheet in ("Питание", "Достопримечательности"):
            formulas["sum_group"] = sumif("sum_group")
            formulas["sum_extra"] = sumif("sum_extra")
        totals_addr[section] = total_row(ws, cols, max(last, 1) + 2, "Итого по выбранным", formulas)

    fill_budget(ws_budget, data, totals_addr, gl, el, base)
    wb.calculation.fullCalcOnLoad = True
    return wb


def add_reserve_row(ws, cols: list[Col], data: dict, last: int, totals_addr: dict) -> int:
    reserve = data.get("reserve") or {}
    share = td.num(reserve.get("share"), 0.0)
    letters = {col.key: get_column_letter(i) for i, col in enumerate(cols, 1)}
    r = last + 1  # fill_rows возвращает 1 для пустого листа, поэтому резерв встаёт во 2-ю строку
    parts_g = [f"Отели!{totals_addr['hotels']['rub_group']}", f"Транспорт!{totals_addr['transport']['rub_group']}",
               f"Питание!{totals_addr['meals']['rub_group']}", f"Достопримечательности!{totals_addr['attractions']['rub_group']}"]
    parts_e = [f"Отели!{totals_addr['hotels']['rub_extra']}", f"Транспорт!{totals_addr['transport']['rub_extra']}",
               f"Питание!{totals_addr['meals']['rub_extra']}", f"Достопримечательности!{totals_addr['attractions']['rub_extra']}"]
    included = ["проживание", "транспорт", "питание", "билеты"]
    if reserve.get("include_flights"):
        parts_g.insert(0, f"Перелеты!{totals_addr['flights']['rub_group']}")
        parts_e.insert(0, f"Перелеты!{totals_addr['flights']['rub_extra']}")
        included.insert(0, "перелёты")
    for offset, item in enumerate(data["other"]):
        if item.get("in_reserve") in (True, td.YES):
            row = 2 + offset
            parts_g.append(f"{letters['rub_group']}{row}")
            parts_e.append(f"{letters['rub_extra']}{row}")
            included.append(str(item.get("category", "")).lower())
    percent = f"{share * 100:g}%"
    ws.cell(row=r, column=1, value="Резерв")
    ws.cell(row=r, column=2, value=f"{percent} от: {', '.join(included)}")
    ws.cell(row=r, column=3, value=data["trip"]["base_currency"])
    ws[f"{letters['price']}{r}"] = share
    ws[f"{letters['price']}{r}"].number_format = "0%"
    ws[f"{letters['qty_group']}{r}"] = 1
    ws[f"{letters['qty_extra']}{r}"] = 1
    ws[f"{letters['sum_group']}{r}"] = f"={letters['price']}{r}*({'+'.join(parts_g)})"
    ws[f"{letters['sum_extra']}{r}"] = f"={letters['price']}{r}*({'+'.join(parts_e)})"
    ws[f"{letters['rub_group']}{r}"] = f'=IF({letters["selected"]}{r}="Да",{letters["sum_group"]}{r},0)'
    ws[f"{letters['rub_extra']}{r}"] = f'=IF({letters["selected"]}{r}="Да",{letters["sum_extra"]}{r},0)'
    for key in ("sum_group", "sum_extra", "rub_group", "rub_extra"):
        ws[f"{letters[key]}{r}"].number_format = MONEY
    ws[f"{letters['selected']}{r}"] = td.YES if share > 0 else td.NO
    excluded = "" if reserve.get("include_flights") else " Перелёты не входят."
    ws[f"{letters['comment']}{r}"] = reserve.get("comment") or (
        f"Автоматически: {percent} от выбранных сумм ({', '.join(included)}).{excluded}"
    )
    totals_addr["reserve"] = {"rub_group": f"{letters['rub_group']}{r}", "rub_extra": f"{letters['rub_extra']}{r}"}
    return r


def fill_budget(ws, data: dict, totals_addr: dict, gl: str, el: str, base: str) -> None:
    cols = [
        Col("category", "Категория", 22), Col("item", "Позиция", 32), Col("stage", "Город / этап", 42),
        Col("currency", "Валюта", 8), Col("group", f"Сумма: {gl}", 16), Col("extra", f"Сумма: {el}", 16),
        Col("status", "Статус", 16), Col("comment", "Комментарий", 50),
    ]
    style_header(ws, cols, data["trip"]["extra"]["size"] == 0)
    meta = data.get("budget") or {}
    rows = [
        ("flights", "Перелеты", "Авиабилеты", "Перелеты", ""),
        ("hotels", "Проживание", "Выбранные отели", "Отели", hotel_span(data["hotels"])),
        ("transport", "Транспорт", "Выбранные варианты", "Транспорт", "Все дни и маршруты"),
        ("meals", "Питание", "Обеды, ужины, напитки, перекусы", "Питание", date_span(data["meals"])),
        ("attractions", "Достопримечательности", "Выбранные билеты", "Достопримечательности", "По датам маршрута"),
    ]
    r = 2
    first = r
    for key, category, item, sheet, stage in rows:
        m = meta.get(key, {})
        ws.cell(row=r, column=1, value=m.get("category", category))
        ws.cell(row=r, column=2, value=m.get("item", item))
        ws.cell(row=r, column=3, value=m.get("stage", stage))
        ws.cell(row=r, column=4, value=base)
        ws.cell(row=r, column=5, value=f"={sheet}!{totals_addr[key]['rub_group']}").number_format = MONEY
        ws.cell(row=r, column=6, value=f"={sheet}!{totals_addr[key]['rub_extra']}").number_format = MONEY
        ws.cell(row=r, column=7, value=m.get("status", "Детализировано"))
        ws.cell(row=r, column=8, value=m.get("comment", "Итог меняется по столбцу «Выбран»."))
        r += 1
    spec = sheet_specs(gl, el, 0, 0, base)["Прочее"]
    letters = {col.key: get_column_letter(i) for i, col in enumerate(spec, 1)}
    for offset, item in enumerate(data["other"]):
        row = 2 + offset
        ws.cell(row=r, column=1, value=item.get("category"))
        ws.cell(row=r, column=2, value=item.get("item"))
        ws.cell(row=r, column=3, value=item.get("budget_stage", ""))
        ws.cell(row=r, column=4, value=base)
        ws.cell(row=r, column=5, value=f"=Прочее!{letters['rub_group']}{row}").number_format = MONEY
        ws.cell(row=r, column=6, value=f"=Прочее!{letters['rub_extra']}{row}").number_format = MONEY
        ws.cell(row=r, column=7, value=item.get("budget_status", "Ориентир"))
        ws.cell(row=r, column=8, value=item.get("budget_comment", "Детали на листе «Прочее»."))
        r += 1
    for note in data.get("budget_notes", []):
        ws.cell(row=r, column=1, value=note.get("category"))
        ws.cell(row=r, column=2, value=note.get("item"))
        ws.cell(row=r, column=3, value=note.get("stage"))
        ws.cell(row=r, column=4, value=base)
        if note.get("amount_group") not in (None, ""):
            ws.cell(row=r, column=5, value=td.num(note["amount_group"])).number_format = MONEY
        if note.get("amount_extra") not in (None, ""):
            ws.cell(row=r, column=6, value=td.num(note["amount_extra"])).number_format = MONEY
        ws.cell(row=r, column=7, value=note.get("status", "Открыт"))
        ws.cell(row=r, column=8, value=note.get("comment"))
        r += 1
    reserve = data.get("reserve") or {}
    ws.cell(row=r, column=1, value="Резерв")
    ws.cell(row=r, column=2, value=f"{td.num(reserve.get('share'), 0) * 100:g}% наземных расходов")
    ws.cell(row=r, column=3, value="Вся поездка")
    ws.cell(row=r, column=4, value=base)
    ws.cell(row=r, column=5, value=f"=Прочее!{totals_addr['reserve']['rub_group']}").number_format = MONEY
    ws.cell(row=r, column=6, value=f"=Прочее!{totals_addr['reserve']['rub_extra']}").number_format = MONEY
    ws.cell(row=r, column=7, value=reserve.get("budget_status", "Автоматически"))
    ws.cell(row=r, column=8, value=reserve.get("budget_comment", "Пересчитывается по выбранным вариантам."))
    last = r
    r += 2
    ws.cell(row=r, column=1, value="Итого").font = Font(bold=True)
    for column in (5, 6):
        letter = get_column_letter(column)
        cell = ws.cell(row=r, column=column, value=f"=SUM({letter}{first}:{letter}{last})")
        cell.font = Font(bold=True)
        cell.number_format = MONEY
    for column in range(1, 9):
        ws.cell(row=r, column=column).fill = TOTAL_FILL
    for row in ws.iter_rows(min_row=2, max_row=r):
        for cell in row:
            if cell.column in (2, 3, 8):
                cell.alignment = Alignment(wrap_text=True, vertical="top")


# ---------------------------------------------------------------- verify

def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    return mac if os.path.exists(mac) else None


def recalculated_copy(xlsx: Path, workdir: Path) -> Path | None:
    """Пересчитать формулы копии книги через LibreOffice; исходный файл не трогается."""
    soffice = find_soffice()
    if not soffice:
        return None
    src = workdir / "in.xlsx"
    shutil.copy2(xlsx, src)
    out = workdir / "out"
    out.mkdir(exist_ok=True)
    profile = (workdir / "lo_profile").as_uri()
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless", "--calc",
         "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(src)],
        check=True, capture_output=True, timeout=180,
    )
    result = out / "in.xlsx"
    return result if result.exists() else None


def verify_values(recalculated: Path, data: dict) -> list[str]:
    """Сверить пересчитанную книгу: нет ошибок формул, итог «Бюджета» равен расчёту по данным."""
    expected = td.totals(data)["total"]
    wb = load_workbook(recalculated, data_only=True)
    problems = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("#"):
                    problems.append(f"{ws.title}!{cell.coordinate}: ошибка формулы {cell.value}")
    ws = wb["Бюджет"]
    total_row = next(r for r in range(ws.max_row, 0, -1) if ws.cell(row=r, column=1).value == "Итого")
    got = (ws.cell(row=total_row, column=5).value or 0, ws.cell(row=total_row, column=6).value or 0)
    for label, g, e in (("группа", got[0], expected[0]), ("+1", got[1], expected[1])):
        if abs(g - e) > 0.5:
            problems.append(f"Итог ({label}): Excel {g:,.2f} ≠ расчёт {e:,.2f}")
    return problems


# ---------------------------------------------------------------- main

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data", required=True, help="trip_data.json")
    parser.add_argument("--out", required=True, help="путь к книге .xlsx")
    parser.add_argument("--force", action="store_true", help="перезаписать книгу с ручными правками (только по решению пользователя)")
    parser.add_argument("--dry-run", action="store_true", help="только проверить данные и показать итоги")
    parser.add_argument("--no-verify", action="store_true", help="не пересчитывать книгу через LibreOffice")
    args = parser.parse_args()

    data_path, out = Path(args.data), Path(args.out)
    data = td.load(data_path)
    errors, warnings = td.validate(data)
    for w in warnings:
        print(f"ПРЕДУПРЕЖДЕНИЕ: {w}")
    if errors:
        for e in errors:
            print(f"ОШИБКА: {e}")
        print("Книга не собрана: исправьте trip_data.json.")
        return 2
    print(td.totals_report(data))
    if args.dry_run:
        return 0

    sidecar = out.parent / SIDECAR
    lock_files = [out.parent / f"~${out.name}", out.parent / f".~lock.{out.name}#"]
    if any(p.exists() for p in lock_files):
        print(f"Книга {out.name} открыта в Excel/LibreOffice — закройте её и повторите.")
        return 4
    if out.exists():
        state = json.loads(sidecar.read_text(encoding="utf-8")) if sidecar.exists() else {}
        recorded = state.get("workbooks", {}).get(out.name, {}).get("sha256")
        if recorded != sha256(out) and not args.force:
            print(
                f"Книга {out.name} изменена после последней сборки (или собрана не этим скриптом).\n"
                "Пересборка затрёт ручные правки. Режим работы: выдавать изменения сводкой ячеек.\n"
                "Пересобрать можно только по явному решению пользователя: сначала перенести правки в данные\n"
                "(xlsx_to_data.py), затем запустить с --force."
            )
            return 3
        archive = out.parent / "архив"
        archive.mkdir(exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        backup = archive / f"{out.stem}_{stamp}{out.suffix}"
        shutil.copy2(out, backup)
        print(f"Резервная копия: {backup}")

    wb = build(data)
    with tempfile.TemporaryDirectory() as tmp:
        built = Path(tmp) / "built.xlsx"
        wb.save(built)
        final, status = built, "LibreOffice не найден: формулы посчитает Excel при открытии, сверка пропущена."
        problems: list[str] = []
        if not args.no_verify:
            recalculated = recalculated_copy(built, Path(tmp))
            if recalculated is not None:
                problems = verify_values(recalculated, data)
                if problems:
                    status = "Книга сохранена без посчитанных значений: сверка не прошла."
                else:
                    # Пересохранённая LibreOffice книга содержит посчитанные значения: превью на телефоне не пустое.
                    final = recalculated
                    status = "Проверка: формулы пересчитаны LibreOffice, итоги совпадают с расчётом по данным."
        else:
            status = "Сверка отключена (--no-verify)."
        staging = out.with_name(out.stem + ".tmp" + out.suffix)
        shutil.copyfile(final, staging)
        os.replace(staging, out)

    state = json.loads(sidecar.read_text(encoding="utf-8")) if sidecar.exists() else {}
    state.setdefault("workbooks", {})[out.name] = {
        "sha256": sha256(out),
        "built_at": dt.datetime.now().isoformat(timespec="seconds"),
        "data_file": data_path.name,
        "data_sha256": sha256(data_path),
    }
    sidecar.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Книга собрана: {out}")
    for problem in problems:
        print(f"ПРОВЕРКА: {problem}")
    print(status)
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
