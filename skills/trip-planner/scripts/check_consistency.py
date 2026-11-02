#!/usr/bin/env python3
"""Проверить согласованность данных поездки, Excel и итоговой программы.

Пример:
    bash scripts/run.sh check_consistency.py --data trip_data.json --program Подробная_программа_Италия.md
    # если Excel правится вручную — брать выбор отелей/ресторанов из книги:
    bash scripts/run.sh check_consistency.py --data trip_data.json --xlsx Расчеты_Италия.xlsx --program ...

Что проверяется:
  * данные: ровно один выбранный вариант на ID, покрытие ночей отелями, обед и ужин в каждый полный день;
  * программа: заголовки дней («### 17 октября, суббота — ...») и верные дни недели, все дни на месте,
    строки «Ночёвка» и «Пешеходная дистанция», выбранный отель в «Ночёвке»;
  * выбранные рестораны упомянуты в своём дне, а основные «обед:/ужин:» программы выбраны в данных;
  * плановые чеки в программе совпадают с суммами в данных;
  * выбранные билеты/места упомянуты в своём дне (имя или aliases).
Код выхода 1 — есть ошибки; предупреждения требуют ручной проверки.
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tripdata as td  # noqa: E402

HEADING = re.compile(r"^###\s+(\d{1,2})\s+([А-Яа-яёЁ]+),\s*([А-Яа-яёЁ]+)\s*[—–-]\s*(.+?)\s*$")
MAIN_MEAL = re.compile(r"(обед|ужин):\s*\*{0,2}(.+?)(?:\.\*\*|\*\*|\.\s|$)", re.IGNORECASE)
MONEY = re.compile(r"(\d{1,3}(?:[\s  ]\d{3})+|\d+)\s*([A-Z]{3})")


def norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[’'`´]", "", text)
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def mentioned(name: str, text_norm: str, aliases=()) -> bool:
    for candidate in [name, *aliases]:
        for part in re.split(r"\s*/\s*|\s*;\s*", candidate or ""):
            n = norm(part)
            if not n:
                continue
            if f" {n} " in f" {text_norm} ":
                return True
            words = [w for w in n.split() if len(w) >= 3]
            if len(words) >= 2 and all(f" {w} " in f" {text_norm} " for w in words[:2]):
                return True
    return False


def parse_program(path: Path, year: int, start_month: int) -> dict[dt.date, dict]:
    days: dict[dt.date, dict] = {}
    current = None
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("## ") or line.startswith("# "):
            current = None
            continue
        if line.startswith("### "):
            m = HEADING.match(line)
            current = None
            if not m:
                continue
            day_num, month_name, weekday, title = m.groups()
            month_name = month_name.lower()
            if month_name not in td.MONTHS_RU_GEN:
                continue
            month = td.MONTHS_RU_GEN.index(month_name) + 1
            y = year + (1 if month < start_month else 0)
            try:
                date = dt.date(y, month, int(day_num))
            except ValueError:
                continue
            current = {"heading": line, "weekday": weekday.lower(), "title": title, "lines": []}
            days[date] = current
            continue
        if current is not None:
            current["lines"].append(line)
    for info in days.values():
        info["text"] = "\n".join(info["lines"])
        info["norm"] = norm(info["text"])
    return days


def amounts_near(lines: list[str], venue: str, currency: str) -> set[int]:
    """Суммы в валюте рядом со всеми упоминаниями ресторана в дне (строка упоминания и следующая)."""
    found: set[int] = set()
    for i, line in enumerate(lines):
        if mentioned(venue, norm(line)):
            for chunk in lines[i:i + 2]:
                for number, cur in MONEY.findall(chunk):
                    if cur == currency:
                        found.add(int(re.sub(r"\D", "", number)))
    return found


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data", required=True)
    parser.add_argument("--program", help="Подробная_программа_*.md")
    parser.add_argument("--xlsx", help="взять разделы из книги вместо trip_data.json")
    args = parser.parse_args()

    data = td.load(args.data)
    if args.xlsx:
        import xlsx_to_data  # noqa: WPS433 — тяжёлый импорт только по запросу

        from_book = xlsx_to_data.extract(Path(args.xlsx), data["trip"])
        # aliases живут только в trip_data.json — переносим их на строки книги
        for section, key in (("meals", "venue"), ("attractions", "name")):
            aliases = {(i.get("id"), norm(i.get(key, ""))): i["aliases"] for i in data[section] if i.get("aliases")}
            for item in from_book[section]:
                found = aliases.get((item.get("id"), norm(item.get(key, ""))))
                if found:
                    item["aliases"] = found
        for key in td.SECTIONS + ("rates", "reserve"):
            data[key] = from_book[key]
        td.apply_defaults(data)

    errors, warnings = td.validate(data)
    info: list[str] = []
    gs = td.group_size(data)
    es = td.extra_size(data)

    for section in td.ID_SECTIONS:
        ids: dict[str, int] = {}
        for item in data[section]:
            ids.setdefault(item["id"], 0)
            ids[item["id"]] += td.selected(item)
        empty = sorted(k for k, v in ids.items() if v == 0)
        if empty:
            info.append(f"{section}: без выбранного варианта (норма для опций): {', '.join(empty)}")

    meals_by_day: dict[dt.date, list[dict]] = {}
    for meal in data["meals"]:
        if td.selected(meal) and meal.get("date"):
            meals_by_day.setdefault(td.parse_date(meal["date"]), []).append(meal)
    nights = td.hotel_nights(data, "group")
    if nights:
        first, last = min(nights), max(nights) + dt.timedelta(days=1)
        day = first + dt.timedelta(days=1)
        while day < last:
            kinds = {norm(m.get("meal", "")) for m in meals_by_day.get(day, [])}
            for need in ("обед", "ужин"):
                if need not in kinds:
                    warnings.append(f"{day}: не выбран {need} (полный день между заездом и выездом)")
            day += dt.timedelta(days=1)

    if args.program:
        start = td.parse_date(data["trip"].get("start_date")) or min(nights or [dt.date.today()])
        program = parse_program(Path(args.program), start.year, start.month)
        if not program:
            errors.append("В программе не найдено ни одного заголовка дня вида «### 17 октября, суббота — ...»")
        for date, day_info in sorted(program.items()):
            real = td.WEEKDAYS_RU[date.weekday()]
            if day_info["weekday"] != real:
                errors.append(f"{day_info['heading']}: {date:%d.%m.%Y} — это {real}, а не {day_info['weekday']}")
            if "Пешеходная дистанция" not in day_info["text"]:
                warnings.append(f"{date}: нет строки «Пешеходная дистанция»")
            if "Ночёвка" not in day_info["text"] and "Ночевка" not in day_info["text"]:
                warnings.append(f"{date}: нет строки «Ночёвка»")
        if nights:
            day = min(nights)
            while day <= max(nights) + dt.timedelta(days=1):
                if day not in program:
                    errors.append(f"{day}: в программе нет дня (поездка с отелями {min(nights)}–{max(nights) + dt.timedelta(days=1)})")
                day += dt.timedelta(days=1)
        for date, hotel_names in nights.items():
            if date not in program:
                continue
            overnight = next((ln for ln in program[date]["lines"] if "Ночёвка" in ln or "Ночевка" in ln), "")
            for name in hotel_names:
                if overnight and not mentioned(name, norm(overnight)):
                    warnings.append(f"{date}: в «Ночёвке» не указан выбранный отель {name}")

        for date, meals in sorted(meals_by_day.items()):
            if date not in program:
                continue
            day_info = program[date]
            for meal in meals:
                venue = meal.get("venue", "")
                kind = norm(meal.get("meal", ""))
                if kind not in ("обед", "ужин"):
                    continue
                if not mentioned(venue, day_info["norm"], meal.get("aliases", ())):
                    errors.append(f"{date}: выбранный {kind} «{venue}» ({meal.get('id')}) не упомянут в программе дня")
                    continue
                cur = meal.get("currency") or data["trip"]["base_currency"]
                found = amounts_near(day_info["lines"], venue, cur)
                if not found:
                    continue
                g, e = td.row_amounts("meals", meal, data)
                food = td.num(meal.get("food_pp"))
                expected = {int(round(food)), int(round(food * gs))} if kind == "обед" else {int(round(g))}
                if kind == "ужин" and es:
                    expected_extra = int(round(e))
                    if expected_extra not in found and len(found) > 1:
                        warnings.append(f"{date}: ужин «{venue}» — в программе {sorted(found)} {cur}, в данных для {data['trip']['extra']['label']}: {expected_extra}")
                if not expected & found:
                    warnings.append(f"{date}: {kind} «{venue}» — в программе {sorted(found)} {cur}, в данных {sorted(expected)}")
            for line in day_info["lines"]:
                if "Запасн" in line or "запасн" in line:
                    continue
                for kind, name in MAIN_MEAL.findall(line):
                    name = name.strip(" *.")
                    if not name or len(name) > 80:
                        continue
                    if not any(mentioned(m.get("venue", ""), norm(name), m.get("aliases", ())) for m in meals):
                        errors.append(f"{date}: в программе {kind.lower()} «{name}», но в данных на этот день выбран другой ресторан")

        for item in data["attractions"]:
            if not td.selected(item) or not item.get("date"):
                continue
            date = td.parse_date(item["date"])
            if date in program and not mentioned(item.get("name", ""), program[date]["norm"], item.get("aliases", ())):
                warnings.append(f"{date}: выбранное место «{item.get('name')}» ({item.get('id')}) не найдено в программе дня — проверить или добавить aliases")

    print("Проверка согласованности")
    print(f"ОШИБКИ: {len(errors)}")
    for e in errors:
        print(f"  - {e}")
    print(f"ПРЕДУПРЕЖДЕНИЯ: {len(warnings)}")
    for w in warnings:
        print(f"  - {w}")
    if info:
        print("ИНФОРМАЦИЯ:")
        for i in info:
            print(f"  - {i}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
