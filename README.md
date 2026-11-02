# my-agent-skills

Личные skills для Claude Code.

| Skill | Что делает |
|---|---|
| [`trip-planner`](skills/trip-planner/SKILL.md) | Планирование поездки: опрос по достопримечательностям → каркас маршрута → детали из внешних источников (курс ЦБ, Островок, официальные сайты) → бюджет в Excel → подробная программа по дням → проверка согласованности. |

## Установка на новом компьютере

```bash
git clone git@github.com:Neraverin/my-agent-skills.git ~/sources/my-agent-skills
cd ~/sources/my-agent-skills
./install.sh
```

`install.sh` создаёт симлинки `~/.claude/skills/<skill>` → `skills/<skill>` (после `git pull` установленный skill сразу обновлён) и готовит окружение скриптов. Параметры: `--copy` — копировать вместо симлинка, `--target DIR` — другой каталог skills, `--no-deps` — без подготовки окружения, имена skills — установить только их. Если skill с таким именем уже установлен, прежняя версия переносится в `~/.claude/skills/.backup/`.

Новая сессия Claude Code подхватывает skills автоматически; проверить можно командой `/trip-planner` или фразой «планируем поездку в …».

## Зависимости trip-planner

| Что | Зачем | Как поставить |
|---|---|---|
| Python 3.10+ с модулем `venv` | Скрипты (сборка Excel, проверки) | Debian/Ubuntu: `sudo apt install python3 python3-venv`; macOS: `brew install python` |
| `openpyxl` | Работа с Excel | Ставится автоматически в `~/.cache/trip-planner/venv` из `skills/trip-planner/scripts/requirements.txt` при первом запуске `run.sh` |
| LibreOffice (необязательно) | Пересчёт и сверка формул собранной книги, посчитанные значения в файле | Debian/Ubuntu: `sudo apt install libreoffice-calc`; macOS: `brew install --cask libreoffice` |
| Доступ в интернет | Курс ЦБ РФ (cbr.ru), Островок, OpenStreetMap (расстояния) | — |

Проверка окружения: `bash ~/.claude/skills/trip-planner/scripts/run.sh --check`.

## Структура

```
skills/
  trip-planner/
    SKILL.md          — описание процесса (фазы, правила, работа с Excel)
    references/       — подробности по фазам, схема данных, формат программы, пример Японии 2026
    assets/           — шаблоны: AGENTS.md поездки, рабочий документ, опрос, программа, визовая программа, пример trip_data.json
    scripts/          — сборка и чтение Excel, проверка согласованности, курс ЦБ, Островок, пешие расстояния
install.sh            — установка skills в ~/.claude/skills
```
