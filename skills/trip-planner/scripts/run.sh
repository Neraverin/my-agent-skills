#!/usr/bin/env bash
# Запуск скриптов trip-planner в отдельном venv с зависимостями из requirements.txt.
# Пример: bash ~/.claude/skills/trip-planner/scripts/run.sh build_workbook.py --data trip_data.json --out Расчеты_X.xlsx
#         bash ~/.claude/skills/trip-planner/scripts/run.sh --check
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="${TRIP_PLANNER_VENV:-$HOME/.cache/trip-planner/venv}"
REQUIREMENTS="$DIR/requirements.txt"

ensure_venv() {
  if [ ! -x "$VENV/bin/python" ]; then
    mkdir -p "$(dirname "$VENV")"
    if ! python3 -m venv "$VENV" >&2; then
      rm -rf "$VENV"
      echo "Не удалось создать venv. Debian/Ubuntu: sudo apt install python3-venv; macOS: brew install python" >&2
      exit 70
    fi
  fi
  # Ставим зависимости, если их нет или requirements.txt изменился после прошлой установки.
  local stamp="$VENV/.requirements.sha256"
  local want
  want="$(sha256sum "$REQUIREMENTS" 2>/dev/null || shasum -a 256 "$REQUIREMENTS")"
  want="${want%% *}"
  if [ ! -f "$stamp" ] || [ "$(cat "$stamp")" != "$want" ] || ! "$VENV/bin/python" -c "import openpyxl" 2>/dev/null; then
    "$VENV/bin/pip" install -q -r "$REQUIREMENTS" >&2
    echo "$want" > "$stamp"
  fi
}

if [ "${1:-}" = "--check" ]; then
  ensure_venv
  echo "python: $("$VENV/bin/python" --version 2>&1) ($VENV)"
  echo "openpyxl: $("$VENV/bin/python" -c 'import openpyxl; print(openpyxl.__version__)')"
  if command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1 \
     || [ -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]; then
    echo "LibreOffice: есть (пересчёт и сверка формул книги)"
  else
    echo "LibreOffice: нет — книга собирается, но формулы не сверяются и посчитаются только при открытии в Excel"
  fi
  exit 0
fi

if [ $# -lt 1 ]; then
  echo "Использование: run.sh <script.py> [аргументы] | run.sh --check" >&2
  exit 64
fi
script="$1"; shift
ensure_venv
exec "$VENV/bin/python" -I -B "$DIR/$script" "$@"
