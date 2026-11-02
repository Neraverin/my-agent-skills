#!/usr/bin/env bash
# Установка skills из этого репозитория в Claude Code.
#
#   ./install.sh                 # все skills, симлинками в ~/.claude/skills
#   ./install.sh trip-planner    # только указанные
#   ./install.sh --copy          # копией вместо симлинка (если репозиторий не будет лежать на месте)
#   ./install.sh --target DIR    # другой каталог skills
#   ./install.sh --no-deps       # не готовить окружение скриптов
#
# Симлинк удобнее: git pull сразу обновляет установленный skill.
# Если в каталоге уже есть skill с тем же именем (не наш симлинк), он переносится в <target>/.backup/.
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
MODE="link"
DEPS=1
NAMES=()

while [ $# -gt 0 ]; do
  case "$1" in
    --copy) MODE="copy" ;;
    --target) TARGET="$2"; shift ;;
    --no-deps) DEPS=0 ;;
    -h|--help) sed -n '2,12p' "$0"; exit 0 ;;
    -*) echo "Неизвестный параметр: $1" >&2; exit 64 ;;
    *) NAMES+=("$1") ;;
  esac
  shift
done

if [ ${#NAMES[@]} -eq 0 ]; then
  for dir in "$REPO"/skills/*/; do
    [ -f "$dir/SKILL.md" ] && NAMES+=("$(basename "$dir")")
  done
fi

mkdir -p "$TARGET"
stamp="$(date +%Y-%m-%d_%H%M%S)"

for name in "${NAMES[@]}"; do
  src="$REPO/skills/$name"
  dst="$TARGET/$name"
  if [ ! -f "$src/SKILL.md" ]; then
    echo "Нет skill: $name" >&2
    exit 66
  fi
  if [ -L "$dst" ] && [ "$(readlink "$dst")" = "$src" ] && [ "$MODE" = "link" ]; then
    echo "✓ $name: уже установлен ($dst → $src)"
  else
    if [ -e "$dst" ] || [ -L "$dst" ]; then
      mkdir -p "$TARGET/.backup"
      mv "$dst" "$TARGET/.backup/$name-$stamp"
      echo "  прежняя версия $name перенесена в $TARGET/.backup/$name-$stamp"
    fi
    if [ "$MODE" = "link" ]; then
      ln -s "$src" "$dst"
      echo "✓ $name: симлинк $dst → $src"
    else
      cp -R "$src" "$dst"
      find "$dst" -name "__pycache__" -prune -exec rm -rf {} +
      echo "✓ $name: скопирован в $dst"
    fi
  fi
  if [ "$DEPS" = 1 ] && [ -x "$src/scripts/run.sh" ]; then
    echo "  окружение скриптов $name:"
    bash "$dst/scripts/run.sh" --check | sed 's/^/    /'
  fi
done

echo "Готово. Skills подхватываются новой сессией Claude Code."
