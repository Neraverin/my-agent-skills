#!/usr/bin/env python3
"""Показать ячейки книги: формулы и пересчитанные значения. Книга не изменяется.

Нужен, когда Excel правят вручную и изменения выдаются сводкой ячеек:
скрипт даёт точные адреса и текущие значения.

Примеры:
    bash scripts/run.sh inspect_workbook.py --xlsx Расчеты_Италия.xlsx --sheet Питание --find "Trattoria"
    bash scripts/run.sh inspect_workbook.py --xlsx Расчеты_Италия.xlsx --sheet Бюджет
    bash scripts/run.sh inspect_workbook.py --xlsx Расчеты_Италия.xlsx --sheet Отели --rows 2-6
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_workbook import recalculated_copy  # noqa: E402

from openpyxl import load_workbook  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--xlsx", required=True)
    parser.add_argument("--sheet", action="append", help="лист; по умолчанию все")
    parser.add_argument("--find", help="показать только строки, где есть этот текст")
    parser.add_argument("--rows", help="диапазон строк, например 2-10")
    parser.add_argument("--no-recalc", action="store_true", help="не пересчитывать через LibreOffice, брать кэш")
    args = parser.parse_args()

    xlsx = Path(args.xlsx)
    wb_f = load_workbook(xlsx, data_only=False)
    with tempfile.TemporaryDirectory() as tmp:
        copy = None if args.no_recalc else recalculated_copy(xlsx, Path(tmp))
        wb_v = load_workbook(copy or xlsx, data_only=True)
        if copy is None and not args.no_recalc:
            print("(LibreOffice не найден: значения формул из кэша файла, могут отсутствовать)")
        lo, hi = 1, 10**9
        if args.rows:
            a, _, b = args.rows.partition("-")
            lo, hi = int(a), int(b or a)
        needle = (args.find or "").lower()
        for ws in wb_f.worksheets:
            if args.sheet and ws.title not in args.sheet:
                continue
            wv = wb_v[ws.title]
            headers = {c.column_letter: c.value for c in ws[1]}
            print(f"=== {ws.title}")
            for row in ws.iter_rows(min_row=max(lo, 1), max_row=min(hi, ws.max_row)):
                cells = [c for c in row if c.value is not None]
                if not cells:
                    continue
                if needle and not any(needle in str(c.value).lower() for c in cells):
                    continue
                for c in cells:
                    value = wv[c.coordinate].value
                    head = headers.get(c.column_letter) if c.row > 1 else None
                    label = f" [{head}]" if head else ""
                    if isinstance(c.value, str) and c.value.startswith("="):
                        print(f"  {c.coordinate}{label}: {c.value}  → {value!r}")
                    else:
                        print(f"  {c.coordinate}{label}: {c.value!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
