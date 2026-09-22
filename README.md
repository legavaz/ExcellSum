# Subtotals API

HTTP-сервис: принимает `.xlsx`-выгрузку из 1С и заменяет статические итоги
группировок на формулы Excel (`=SUM(...)` / ссылки на ячейки).

- уровни строк: колонка с уровнем → сохранённые группировки Excel (outline) → отступ ячейки;
- итоги сверху, снизу или смешанно (`style=auto|top|bottom`);
- многоуровневая вложенность: внешний итог ссылается на итоги подгрупп, а не на детали,
  поэтому двойного суммирования не возникает;
- общий итог (последняя строка) собирается по итогам групп верхнего уровня
  (`=H7+H697+...`), а не по всем числовым строкам подряд;
- формулы: один ребёнок → `=C5`, непрерывный диапазон → `=SUM(C2:C5)`,
  диапазон с пропусками → `=C2+C5+C9`;
- детальные строки и колонки вне `columns` не изменяются, исходный файл не перезаписывается.

## Быстрый старт

```bash
docker compose up --build -d
curl http://localhost:8000/health
```

Веб-форма: http://localhost:8000/ — выбор `.xlsx`, параметры, кнопка «Обработать»
и ссылка на скачивание готового файла (при dry-run показывается лог).
Колонки со значениями можно не указывать — определяются автоматически.
При выборе файла форма сама вызывает `/analyze`: для плоского отчёта включает
галочку «Плоский отчёт» и показывает лист, строку итога и заменяемые колонки.

Swagger UI: http://localhost:8000/docs, ReDoc: http://localhost:8000/redoc,
OpenAPI JSON: http://localhost:8000/openapi.json.

## HTTP API

| Метод | Путь | Назначение |
|---|---|---|
| GET | `/health` | проверка живости, `{"status": "ok"}` |
| POST | `/analyze` | разбор `.xlsx`: найденные группировки, режим и колонки значений |
| POST | `/convert` | конвертация `.xlsx` (multipart/form-data) |

Поля `/convert`:

| Поле | Обяз. | Описание |
|---|---|---|
| `file` | да | файл `.xlsx` |
| `columns` | нет | колонки со значениями, напр. `E,F,G`; пусто — автоопределение: берутся колонки, где в итоговой строке число и есть числа у строк-детей |
| `level_column` | нет | колонка с номером уровня, напр. `A` (иначе outline/indent) |
| `style` | нет | `auto` (по умолчанию), `top`, `bottom` |
| `sheet` | нет | имя листа (по умолчанию активный) |
| `dry_run` | нет | `true` — вернуть JSON с логом, файл не создаётся |
| `grand_total` | нет | `true` — плоский отчёт без группировок: последняя числовая строка суммирует непрерывный блок выше |

Авторизация не требуется.

```bash
curl -X POST http://localhost:8000/convert \
  -F "file=@report.xlsx" \
  -F "columns=E,F,G" \
  -F "style=auto" \
  -o report_fix.xlsx

curl -X POST http://localhost:8000/convert \
  -F "file=@report.xlsx" -F "columns=E,F,G" -F "dry_run=true"
```

Ответ 200 при обычном запуске — сам `.xlsx`; заголовки `X-Job-Id` и
`X-Log-Preview` (последние 1.5 КБ лога). Ошибки: 400 — параметры/формат,
500 — сбой конвертации.

## Переменные окружения

| Переменная | По умолчанию | Описание |
|---|---|---|
| `API_PORT` | `8000` | внешний порт docker compose (если 8000 занят: `API_PORT=8001 docker compose up -d`) |
| `STORAGE_DIR` | `/data/storage` | каталог загрузок и результатов |
| `STORAGE_TTL_HOURS` | `24` | TTL файлов, фоновая очистка каждые 10 минут |

## CLI (то же ядро без HTTP)

```bash
python subtotals.py report.xlsx -o report_fix.xlsx
python subtotals.py report.xlsx -c "E,F,G" -o report_fix.xlsx
python subtotals.py report.xlsx -c "E,F,G" -l A --style bottom -o report_fix.xlsx
python subtotals.py report.xlsx -c "E,F,G" --dry-run
python subtotals.py report.xlsx -c "D,H" --grand-total -o report_fix.xlsx
```

## Тесты

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements-dev.txt
.venv\Scripts\pytest
```

В контейнере (образ ставит `requirements-dev.txt` и копирует `tests/`):

```bash
docker compose run --rm api pytest -q
```

Выгрузка OpenAPI:

```bash
python export_openapi.py openapi.json
# или
docker compose run --rm api python export_openapi.py /data/openapi.json
```

## Структура

```
.
├── subtotals.py          # ядро конвертации + CLI
├── app.py                # FastAPI, /, /health и /convert
├── storage.py            # каталоги заданий + TTL-очистка
├── schemas.py            # Pydantic-схемы ответов
├── openapi_meta.py       # теги и описание Swagger
├── export_openapi.py     # выгрузка openapi.json
├── web/index.html        # веб-форма
├── requirements.txt      # рантайм-зависимости
├── requirements-dev.txt  # рантайм + pytest/httpx
├── Dockerfile
├── docker-compose.yml
├── pytest.ini
├── docs/SPEC.md          # исходная спека из диалога
└── tests/
    ├── conftest.py       # env до импорта app
    ├── test_core.py      # golden-кейсы T1-T6
    └── test_api.py       # HTTP-контракт
```

## Отличия от исходной спеки

Исправлено при реализации (детали — в `docs/SPEC.md`):

1. `cleanup_loop` сделан `async` — в спеке синхронная функция передавалась
   в `asyncio.create_task()`, из-за чего сервис не стартовал.
2. `dry_run` действительно не создаёт файл: флаг прокинут в `process()`.
3. `X-Log-Preview` санитизируется до latin-1 — кириллица в логе больше не
   ломает заголовок ответа.
4. Dockerfile копирует `tests/` и ставит `requirements-dev.txt`, поэтому
   `docker compose run --rm api pytest -q` из спеки работает.
5. Golden-тест T5 приведён в соответствие с правилом B3: для дерева
   `A → (x, A1 → y, z)` внешний итог равен `=C2+C3+C5` (итог подгруппы
   учитывается как один атом).
6. Неизвестный лист и пустой `columns` возвращают 400 вместо 500.
7. Авто-режим `style=auto` переписан на разбор дерева группировок
   (`build_tree`/`plan_auto`): эталонный классификатор по соседям ошибался
   в двух golden-кейсах — не находил итог подгруппы в середине отчёта (T3)
   и принимал деталь после подгруппы за нижний итог (T5); кроме того,
   строки-заголовки групп больше не получают формул (T1, T4). Для
   принудительных `style=top|bottom` используется алгоритм из спеки.
   В неоднозначных смешанных структурах авто-режим опирается на преобладание
   стиля в отчёте — при сомнениях задавайте `style` явно.
8. Общий итог в конце таблицы определяется по значению: если он совпадает
   с суммой итогов групп верхнего уровня — ссылается на них, если только с
   суммой одной группы — на неё. Так «Итого» не задваивает вложенные итоги
   (проверено на `пример.xlsx`: 85 и 256 итогов групп, обе суммы = 1423).
