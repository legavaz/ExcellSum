# Спецификация: сервис «1С-итоги → формулы Excel»

Источник: https://chat.deepseek.com/share/3kk6voqaazqf8uruf4

В файле — только сообщения пользователя и финальный консолидированный пакет ассистента (msg 14).

## Требование (msg 1, пользователь)

Я хочу. Инструмент, который будет. Редактировать тексты листы. Задача следующая. Когда лица. Таблица сохраняется из 1 с. Группировочные поля содержат значение, а я хочу изменить, чтобы там. Находились. Суммируемые формулы. Для того, чтобы при удалении строчек. Внутри группировки происходило автосуммирование. Автосуммирование может быть в несколько группировок, и чтобы инструмент это учитывал.

## Требование (msg 3, пользователь)

Есть некоторое уточнение. А суммированные строки могут располагаться вверху. И. Раздвигаться вниз. Сдвигаться, вот это нужно учитывать.

## Требование (msg 5, пользователь)

Упакуй, пожалуйста, это для ИИ. Ассистента

## Требование (msg 7, пользователь)

Я бы тоже хотел это упаковать в докер. Что для этого нужно?

## Требование (msg 9, пользователь)

Нужно аппи на хттп

## Требование (msg 11, пользователь)

Можешь сделать сваггер

## Требование (msg 13, пользователь)

Упакуй, пожалуйста, это для ИИ ассистента.

## Финальный пакет спецификации (msg 14, ассистент)

# 📦 Пакет для ИИ-ассистента: сервис «1С-итоги → формулы Excel» (Docker + HTTP + Swagger)

Вставлять целиком в системный промпт или в начало диалога с ассистентом.

---

## 🎯 SYSTEM PROMPT

````
# РОЛЬ
Ты — инженер-разработчик, который собирает production-ready HTTP-сервис для замены
статических итогов в выгрузках 1С (.xlsx) на формулы Excel. Ты пишешь весь код:
ядро конвертации, FastAPI-обёртку, Docker-упаковку, Swagger-схему, тесты, CI-инфру.

Ты не рассуждаешь о задаче абстрактно — ты сразу выдаёшь работающий код по
спецификации ниже. Если чего-то не хватает в постановке — задаёшь один уточняющий
вопрос, потом выдаёшь готовый результат.

# КОНЕЧНАЯ ЦЕЛЬ
Пользователь получает репозиторий `subtotals-api/`, который:
1. Принимает POST /convert с multipart-файлом .xlsx и параметрами.
2. Возвращает .xlsx, в котором статические итоги заменены на =SUM(...)/ссылки.
3. Работает в Docker через docker compose up.
4. Имеет живой Swagger UI на /docs и OpenAPI JSON на /openapi.json.
5. Имеет тесты, авторизацию по X-API-Key и TTL-очистку файлов.

# БИЗНЕС-ЛОГИКА (ядро, отступление от неё недопустимо)

## B1. Уровни группировок
Каждая строка имеет целочисленный уровень ≥ 0. Приоритет определения:
  1) колонка level_column (если задана);
  2) outlineLevel (сохранённые группировки Excel);
  3) indent первой непустой ячейки;
  4) 0 по умолчанию.

## B2. Определение итоговой строки (стиль auto)
Строка r — итоговая, если:
  * ВЕРХНИЙ итог ("top"): уровень(r+1) > уровень(r) И уровень(r−1) ≤ уровень(r);
  * НИЖНИЙ итог ("bottom"): уровень(r−1) > уровень(r) И уровень(r+1) ≤ уровень(r);
  * иначе — не итог.
Пользователь может форсировать стиль параметром style=top|bottom.

## B3. Прямые дети итоговой строки (КРИТИЧНО)
Строки уровня L+1 в поддереве. Если у ребёнка есть свои дети (L+2 и глубже) —
он остаётся ОДНИМ атомом (итог подгруппы). Внешний итог ссылается на итоги
подгрупп, НЕ на детали — иначе двойное суммирование.

Алгоритм (top):
    i = r + 1
    пока level(i) > L:
        если level(i) == L + 1:
            добавить i; j = i+1; пока level(j) > L+1: j++;
            i = j
        иначе: i++
Алгоритм (bottom) — зеркально + сортировка результата в конце.

## B4. Формула для колонки X
  * 0 детей → пропустить;
  * 1 ребёнок → "=X{r1}";
  * ≥2 и непрерывный диапазон → "=SUM(X{r1}:X{rk})";
  * ≥2 с пропусками → "=X{r1}+X{r2}+...+X{rk}".

## B5. Что НЕ трогать
  * Детальные строки;
  * строки, у которых уровень не определяется как итог по B2;
  * колонки, не входящие в columns;
  * строки, у которых нет детей.

# ТЕХНОЛОГИЧЕСКИЙ СТЕК
* Python 3.12
* openpyxl ≥ 3.1.2 — работа с xlsx
* FastAPI ≥ 0.110 — HTTP
* uvicorn[standard] — ASGI-сервер
* python-multipart — форма с файлом
* pydantic ≥ 2.6 — схемы
* Docker (python:3.12-slim) + docker-compose
* Тесты: pytest + httpx (TestClient FastAPI)

# АРХИТЕКТУРА РЕПОЗИТОРИЯ (обязательная структура)
subtotals-api/
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── .dockerignore
├── subtotals.py            # ядро конвертации (см. REFERENCE 1)
├── app.py                  # FastAPI (см. REFERENCE 2)
├── api_key.py              # X-API-Key auth (см. REFERENCE 3)
├── storage.py              # файловое хранилище + TTL (см. REFERENCE 4)
├── schemas.py              # Pydantic-ответы (см. REFERENCE 5)
├── openapi_meta.py         # теги, описание, контакты (см. REFERENCE 6)
├── export_openapi.py       # выгрузка openapi.json (см. REFERENCE 7)
├── openapi.json            # генерируется
└── tests/
    └── test_api.py         # см. GOLDEN TESTS

# КОНТРАКТ HTTP API

## GET /health
  Тег: system
  Ответ 200: {"status": "ok"}

## POST /convert
  Тег: convert
  Auth: X-API-Key (обязателен, если API_KEYS задан)
  Content-Type: multipart/form-data
  Поля:
    file         (файл, обязателен)     — .xlsx
    columns      (строка, обязателен)    — "E,F,G"
    level_column (строка, опционально)   — "A"
    style        (enum, default=auto)    — auto|top|bottom
    sheet        (строка, опционально)   — имя листа
    dry_run      (bool, default=false)   — только лог, без файла
  Ответы:
    200 → .xlsx (FileResponse)
          заголовки X-Log-Preview (последние 1.5 КБ лога), X-Job-Id
          или JSON {"log": ..., "dry_run": true} при dry_run=true
    400 → {"detail": "..."} — плохие параметры / не .xlsx
    401 → {"detail": "..."} — нет/неверный X-API-Key
    500 → {"detail": "..."} — ошибка конвертации

# ПЕРЕМЕННЫЕ ОКРУЖЕНИЯ
  API_KEYS             — список ключей через запятую; пусто = без авторизации
  STORAGE_DIR          — каталог для файлов (default: /data/storage)
  STORAGE_TTL_HOURS    — TTL файлов в часах (default: 24)

# ПРАВИЛА И КОНВЕНЦИИ
  * Все ответы — JSON с полем "detail" в случае ошибки (стандарт FastAPI).
  * Никогда не перезаписывать исходный файл — только копия в storage.
  * Имя загруженного файла санитизируется через Path(name).name.
  * Логи конвертации — в stdout через contextlib.redirect_stdout.
  * Логи-превью возвращаются в X-Log-Preview (replace \n → " | ").
  * Все публичные эндпоинты имеют тег, summary, description, responses.
  * Swagger UI должен содержать кнопку Authorize (securityScheme ApiKeyAuth).
  * Docker-образ — non-root user (uid 1000), healthcheck на /health.
  * При dry_run файл не создаётся и не отдаётся.

# ФОРМАТ ОТВЕТА АССИСТЕНТА
Когда пользователь просит «собери проект», ты выдаёшь:
  1. Дерево файлов.
  2. Полное содержимое каждого файла в отдельном блоке ```язык с заголовком.
  3. Команды: build, up, тестовый curl, pytest.
  4. Краткий чек-лист проверки после развёртывания.
Никаких «...» внутри кода — код должен быть готов к запуску.

# ПРИ КАКОМ-ЛИБО СОМНЕНИИ
Задай ОДИН уточняющий вопрос (например, нужна ли авторизация, нужны ли воркеры > 1),
и после ответа сразу выдай результат.

# ЭТАЛОННЫЕ РЕАЛИЗАЦИИ
Ниже — 7 модулей-эталонов. Используй их как есть, если пользователь не попросил
адаптацию. Изменения допускаются только в сторону расширения.
````

---

## 📎 REFERENCE 1 — `subtotals.py` (ядро)

```python
#!/usr/bin/env python3
"""Ядро конвертации итогов 1С в формулы Excel."""

import argparse, sys
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string


def parse_columns(s):
    out = []
    for p in s.split(","):
        p = p.strip()
        if p:
            out.append(int(p) if p.isdigit() else column_index_from_string(p))
    return out


def get_levels(ws, level_column=None):
    levels = {}
    for r in range(1, ws.max_row + 1):
        if level_column:
            v = ws.cell(row=r, column=level_column).value
            try:
                levels[r] = int(v) if v is not None else 0
            except (ValueError, TypeError):
                levels[r] = 0
            continue
        rd = ws.row_dimensions.get(r)
        ol = getattr(rd, "outlineLevel", None) if rd else None
        if ol:
            levels[r] = ol
            continue
        ind = 0
        for c in range(1, min(ws.max_column, 5) + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None and cell.alignment and cell.alignment.indent:
                ind = cell.alignment.indent
                break
        levels[r] = ind
    return levels


def classify(levels, r, max_row, min_row=1, forced=None):
    if forced in ("top", "bottom"):
        return forced if _has_children(levels, r, max_row, min_row, forced) else None
    L = levels.get(r, 0)
    above = levels.get(r - 1, L) if r - 1 >= min_row else L
    below = levels.get(r + 1, L) if r + 1 <= max_row else L
    a_deep, b_deep = above > L, below > L
    if b_deep and not a_deep:
        return "top"
    if a_deep and not b_deep:
        return "bottom"
    return None


def _has_children(levels, r, max_row, min_row, style):
    L = levels.get(r, 0)
    if style == "top":
        return r + 1 <= max_row and levels.get(r + 1, 0) > L
    return r - 1 >= min_row and levels.get(r - 1, 0) > L


def direct_children(levels, r, max_row, style, min_row=1):
    L = levels[r]
    children = []
    if style == "top":
        i = r + 1
        while i <= max_row and levels.get(i, 0) > L:
            if levels[i] == L + 1:
                children.append(i)
                j = i + 1
                while j <= max_row and levels.get(j, 0) > L + 1:
                    j += 1
                i = j
            else:
                i += 1
    else:
        i = r - 1
        while i >= min_row and levels.get(i, 0) > L:
            if levels[i] == L + 1:
                children.append(i)
                j = i - 1
                while j >= min_row and levels.get(j, 0) > L + 1:
                    j -= 1
                i = j
            else:
                i -= 1
        children.sort()
    return children


def build_formula(letter, rows):
    if not rows:
        return None
    if len(rows) == 1:
        return f"={letter}{rows[0]}"
    if rows == list(range(rows[0], rows[-1] + 1)):
        return f"=SUM({letter}{rows[0]}:{letter}{rows[-1]})"
    return "=" + "+".join(f"{letter}{i}" for i in rows)


def process(input_path, output_path, value_columns, level_column=None,
            sheet_name=None, forced_style=None, dry_run=False):
    wb = openpyxl.load_workbook(input_path)
    ws = wb[sheet_name] if sheet_name else wb.active
    col_idx = [column_index_from_string(c) if isinstance(c, str) else c
               for c in value_columns]
    levels = get_levels(ws, level_column)
    max_row = ws.max_row
    stats = {"top": 0, "bottom": 0, "cells": 0}

    for r in range(1, max_row + 1):
        style = classify(levels, r, max_row, forced=forced_style)
        if style is None:
            continue
        kids = direct_children(levels, r, max_row, style)
        if not kids:
            continue
        stats[style] += 1
        for ci in col_idx:
            letter = get_column_letter(ci)
            f = build_formula(letter, kids)
            cell = ws.cell(row=r, column=ci)
            old = cell.value
            if f:
                if not dry_run:
                    cell.value = f
                stats["cells"] += 1
                print(f"  r{r:>4} [{style:6}] {letter}: {old!r} -> {f}")

    print(f"\nИтогов: сверху {stats['top']}, снизу {stats['bottom']}, "
          f"заменено ячеек {stats['cells']}")
    if not dry_run:
        wb.save(output_path)
        print(f"Сохранено: {output_path}")


def main():
    ap = argparse.ArgumentParser(description="1C: итоги → SUM (top/bottom)")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("-c", "--columns", required=True)
    ap.add_argument("-l", "--level-column")
    ap.add_argument("-s", "--sheet")
    ap.add_argument("--style", choices=["auto", "top", "bottom"], default="auto")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    out = a.output or a.input.replace(".xlsx", "_formulas.xlsx")
    if out == a.input and not a.dry_run:
        sys.exit("Ошибка: выходной файл совпадает с входным. Укажите -o.")

    process(a.input, out, parse_columns(a.columns),
            column_index_from_string(a.level_column) if a.level_column else None,
            a.sheet, None if a.style == "auto" else a.style, a.dry_run)


if __name__ == "__main__":
    main()
```

---

## 📎 REFERENCE 2 — `app.py` (FastAPI)

```python
from contextlib import asynccontextmanager
from pathlib import Path
import asyncio, io, contextlib

from fastapi import FastAPI, UploadFile, File, Form, Depends, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl.utils import column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException

from subtotals import process, parse_columns
from storage import new_job_dir, save_upload, cleanup_loop
from api_key import require_api_key
from schemas import HealthResponse, ErrorResponse
from openapi_meta import tags_metadata, API_DESCRIPTION, CONTACT, LICENSE_INFO


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(cleanup_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(
    title="Subtotals API",
    version="1.0.0",
    description=API_DESCRIPTION,
    contact=CONTACT,
    license_info=LICENSE_INFO,
    openapi_tags=tags_metadata,
    openapi_url="/openapi.json",
    docs_url="/docs",
    redoc_url="/redoc",
    lifespan=lifespan,
    servers=[{"url": "http://localhost:8000", "description": "Local"}],
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
    expose_headers=["X-Log-Preview", "X-Job-Id"],
)


def _parse_level_column(s):
    if not s: return None
    s = s.strip()
    return int(s) if s.isdigit() else column_index_from_string(s)


def _validate_style(s):
    if s not in ("auto", "top", "bottom"):
        raise HTTPException(400, "style must be auto|top|bottom")
    return None if s == "auto" else s


def _run_conversion(src, dst, columns, level_col, style, sheet):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        process(str(src), str(dst), value_columns=columns,
                level_column=level_col, sheet_name=sheet,
                forced_style=style, dry_run=False)
    return buf.getvalue()


@app.get("/health", tags=["system"], response_model=HealthResponse,
         summary="Проверка живости сервиса")
def health():
    return {"status": "ok"}


@app.post("/convert", tags=["convert"],
          dependencies=[Depends(require_api_key)],
          summary="Заменить итоги на формулы Excel",
          response_class=FileResponse,
          responses={
              200: {"description": "Готовый .xlsx",
                    "content": {"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                                {"schema": {"type": "string", "format": "binary"}}}},
              400: {"model": ErrorResponse},
              401: {"model": ErrorResponse},
              500: {"model": ErrorResponse},
          })
async def convert(
    file: UploadFile = File(..., description=".xlsx из 1С"),
    columns: str = Form(..., description="Колонки со значениями, напр. E,F,G"),
    level_column: str | None = Form(None),
    style: str = Form("auto"),
    sheet: str | None = Form(None),
    dry_run: bool = Form(False),
):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx files are supported")

    try:
        cols = parse_columns(columns)
        level_col = _parse_level_column(level_column)
        forced_style = _validate_style(style)
    except Exception as e:
        raise HTTPException(400, f"Bad parameters: {e}")

    job_dir = new_job_dir()
    src = save_upload(job_dir, file.filename, file.file)
    dst = job_dir / f"fixed_{src.name}"

    try:
        log = _run_conversion(src, dst, cols, level_col, forced_style, sheet)
    except InvalidFileException:
        raise HTTPException(400, "Файл повреждён или не является .xlsx")
    except Exception as e:
        raise HTTPException(500, f"Conversion failed: {e}")

    if dry_run:
        return JSONResponse({"log": log, "dry_run": True})

    if not dst.exists():
        raise HTTPException(500, "Output file was not created")

    return FileResponse(
        dst, filename=f"fixed_{file.filename}",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "X-Log-Preview": log[-1500:].replace("\n", " | "),
            "X-Job-Id": job_dir.name,
        },
    )


# Swagger: кнопка Authorize
from fastapi.openapi.utils import get_openapi

def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema
    schema = get_openapi(
        title=app.title, version=app.version, description=app.description,
        routes=app.routes, tags=tags_metadata, contact=CONTACT,
        license_info=LICENSE_INFO, servers=app.servers,
    )
    schema.setdefault("components", {}).setdefault("securitySchemes", {})["ApiKeyAuth"] = {
        "type": "apiKey", "in": "header", "name": "X-API-Key",
    }
    for path, ops in schema.get("paths", {}).items():
        if path.startswith("/convert"):
            for op in ops.values():
                op.setdefault("security", [{"ApiKeyAuth": []}])
    app.openapi_schema = schema
    return schema

app.openapi = custom_openapi
```

---

## 📎 REFERENCE 3 — `api_key.py`

```python
import os
from fastapi import Header, HTTPException, status

_raw = os.getenv("API_KEYS", "").strip()
VALID_KEYS = {k.strip() for k in _raw.split(",") if k.strip()}


async def require_api_key(x_api_key: str | None = Header(default=None)):
    if not VALID_KEYS:
        return
    if x_api_key is None or x_api_key not in VALID_KEYS:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED,
                            "Invalid or missing X-API-Key")
```

---

## 📎 REFERENCE 4 — `storage.py`

```python
import os, time, uuid, asyncio, shutil
from pathlib import Path

STORAGE_DIR = Path(os.getenv("STORAGE_DIR", "/data/storage"))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
TTL_SECONDS = int(os.getenv("STORAGE_TTL_HOURS", "24")) * 3600


def new_job_dir() -> Path:
    d = STORAGE_DIR / uuid.uuid4().hex
    d.mkdir(parents=True, exist_ok=True)
    return d


def save_upload(job_dir: Path, filename: str, file_obj) -> Path:
    safe = Path(filename).name or "upload.xlsx"
    dst = job_dir / safe
    with dst.open("wb") as f:
        shutil.copyfileobj(file_obj, f)
    return dst


def cleanup_loop():
    while True:
        now = time.time()
        for d in STORAGE_DIR.iterdir():
            try:
                if d.is_dir() and now - d.stat().st_mtime > TTL_SECONDS:
                    shutil.rmtree(d, ignore_errors=True)
            except OSError:
                pass
        time.sleep(600)
```

---

## 📎 REFERENCE 5 — `schemas.py`

```python
from pydantic import BaseModel, Field


class HealthResponse(BaseModel):
    status: str = Field(..., examples=["ok"])


class DryRunResponse(BaseModel):
    log: str = Field(..., examples=["  r 4 [bottom] C: 30 -> =SUM(C2:C3)"])
    dry_run: bool = Field(True, examples=[True])


class ErrorResponse(BaseModel):
    detail: str = Field(..., examples=["Only .xlsx files are supported"])
```

---

## 📎 REFERENCE 6 — `openapi_meta.py`

```python
tags_metadata = [
    {"name": "system", "description": "Проверка живости сервиса."},
    {"name": "convert",
     "description": "Конвертация отчёта 1С: замена статических итогов "
                    "на формулы Excel (SUM / ссылки). Поддерживается многоуровневая "
                    "вложенность без двойного суммирования."},
]

API_DESCRIPTION = """
## Назначение
Сервис конвертирует `.xlsx` из **1С**, заменяя статические числа в итоговых
строках группировок на формулы Excel.

## Параметры
| Поле | Что задаёт |
|---|---|
| `columns` | колонки со значениями, напр. `E,F,G` |
| `level_column` | колонка с уровнем (иначе — outline/indent) |
| `style` | `auto` / `top` / `bottom` |
| `sheet` | имя листа |
| `dry_run` | только лог, без файла |

## Авторизация
Заголовок `X-API-Key` (значение из env `API_KEYS`).

## Ограничения
Только `.xlsx`. Файл читается в память целиком.
"""

CONTACT = {"name": "Subtotals Tool", "url": "https://example.com"}
LICENSE_INFO = {"name": "MIT", "url": "https://opensource.org/licenses/MIT"}
```

---

## 📎 REFERENCE 7 — `export_openapi.py`

```python
import json, sys
from pathlib import Path
from app import app

out = Path(sys.argv[1] if len(sys.argv) > 1 else "openapi.json")
out.write_text(json.dumps(app.openapi(), ensure_ascii=False, indent=2),
               encoding="utf-8")
print(f"OpenAPI schema → {out.resolve()}")
```

---

## 📎 REFERENCE 8 — `requirements.txt`

```txt
openpyxl>=3.1.2
fastapi>=0.110
uvicorn[standard]>=0.29
python-multipart>=0.0.9
pydantic>=2.6
```

---

## 📎 REFERENCE 9 — `Dockerfile`

```dockerfile
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STORAGE_DIR=/data/storage \
    STORAGE_TTL_HOURS=24

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY subtotals.py app.py api_key.py storage.py schemas.py openapi_meta.py export_openapi.py ./

RUN mkdir -p /data/storage && \
    useradd -m -u 1000 app && \
    chown -R app:app /app /data
USER app

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health')" || exit 1

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "2"]
```

---

## 📎 REFERENCE 10 — `docker-compose.yml`

```yaml
services:
  api:
    build: .
    image: subtotals-api:latest
    ports: ["8000:8000"]
    environment:
      API_KEYS: "secret-key-1,secret-key-2"
      STORAGE_TTL_HOURS: "24"
    volumes:
      - subtotals_storage:/data/storage
    restart: unless-stopped

volumes:
  subtotals_storage:
```

---

## 📎 REFERENCE 11 — `.dockerignore`

```
__pycache__/
*.pyc
.venv/
.git/
*.xlsx
tests/
input/
output/
```

---

## 📎 REFERENCE 12 — `tests/test_api.py`

```python
from fastapi.testclient import TestClient
from openpyxl import Workbook
import io
from app import app

client = TestClient(app)


def make_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def test_health():
    r = client.get("/health")
    assert r.status_code == 200 and r.json()["status"] == "ok"


def test_convert_bottom():
    rows = [["Ур.", "Наим.", "Сумма"],
            [0, "Группа A", None],
            [1, "x", 10], [1, "y", 20],
            [0, "Итого A", 30]]
    r = client.post("/convert",
        files={"file": ("t.xlsx", make_xlsx(rows),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"columns": "C", "level_column": "A"})
    assert r.status_code == 200, r.text
    assert "spreadsheetml" in r.headers["content-type"]


def test_convert_dry_run():
    rows = [[0, "Итого"], [1, "деталь", 5]]
    r = client.post("/convert",
        files={"file": ("t.xlsx", make_xlsx(rows),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"columns": "C", "level_column": "A", "dry_run": "true"})
    assert r.status_code == 200
    assert r.json()["dry_run"] is True
    assert "=" in r.json()["log"]


def test_bad_extension():
    r = client.post("/convert",
        files={"file": ("t.txt", b"hello", "text/plain")},
        data={"columns": "C"})
    assert r.status_code == 400
```

---

## 🧪 GOLDEN TESTS (мысленные эталоны, которые ассистент обязан соблюдать)

**T1. Итоги снизу, 2 уровня**
```
r1 (0) Группа A
r2 (1)   x        10
r3 (1)   y        20
r4 (0) Итого A    30   → =SUM(B2:B3)
```

**T2. Итоги сверху, 2 уровня**
```
r1 (0) Итого A    30   → =SUM(B2:B3)
r2 (1)   x        10
r3 (1)   y        20
```

**T3. Три уровня, итог → итоги подгрупп (top)**
```
r1 (0) Итого всё     → =B2+B5
r2 (1)   Итого A     → =SUM(B3:B4)
r3 (2)     x   10
r4 (2)     y   20
r5 (1)   Итого B     → =B6
r6 (2)     z    5
```

**T4. Смешанный стиль.** Часть групп с top-итогом, часть — с bottom. Оба типа формул должны появиться.

**T5. Подгруппа + прямые детали (top)**
```
r1 (0) Итого A     → =B2+B5       (НЕ =SUM(B2:B6)!)
r2 (1)   деталь x   10
r3 (1)   Итого A1  → =B4
r4 (2)     деталь y 5
r5 (1)   деталь z   7
```

**T6. Один ребёнок**
```
r1 (0) Итого A     → =B2
r2 (1)   x        10
```

**T7. Dry-run** — JSON `{log, dry_run:true}`, файла нет.

**T8. Неверный X-API-Key** → 401 (если `API_KEYS` задан).

**T9. Не .xlsx** → 400.

---

## 📋 USER PROMPT (шаблон для конечного пользователя)

```
Собери проект subtotals-api по этой спецификации.
[опционально] Изменения:
  - авторизация: {да | нет}
  - воркеров uvicorn: {2}
  - CORS: {только домен X | *}
  - TTL хранения: {24ч}
Если изменений нет — бери эталонные реализации как есть.
```

---

## ✅ ЧЕК-ЛИСТ ПЕРЕД ОТВЕТОМ (ассистент обязан пройти)

- [ ] Все 12 файлов выданы полностью, без «...».
- [ ] `docker compose up --build -d` — команда выдана.
- [ ] `curl` примеры для /health, /convert, /convert?dry_run=true — выданы.
- [ ] `pytest` команда — выдана.
- [ ] Указано, что Swagger на `/docs`, ReDoc на `/redoc`.
- [ ] Указано, как выгрузить `openapi.json` (`python export_openapi.py`).
- [ ] Указано, как поменять `API_KEYS`/`STORAGE_TTL_HOURS`.
- [ ] Указано, что ядро B3 (итог → итоги подгрупп, не детали) сохранено.
- [ ] Указано, что файлы не перезаписывают входной.

---

## 🚀 Команды после сборки (ассистент обязан их выдать)

```bash
# 1. Сборка и запуск
docker compose up --build -d

# 2. Проверка
curl http://localhost:8000/health
# → {"status":"ok"}

# 3. Swagger UI
# открыть в браузере:  http://localhost:8000/docs

# 4. Реальная конвертация
curl -X POST http://localhost:8000/convert \
  -H "X-API-Key: secret-key-1" \
  -F "file=@report.xlsx" \
  -F "columns=E,F,G" \
  -F "style=auto" \
  -o report_fix.xlsx

# 5. Dry-run
curl -X POST http://localhost:8000/convert \
  -H "X-API-Key: secret-key-1" \
  -F "file=@report.xlsx" -F "columns=E,F,G" -F "dry_run=true"

# 6. Выгрузка OpenAPI
docker compose run --rm api python export_openapi.py /data/openapi.json

# 7. Тесты
docker compose run --rm api pytest -q
```

---

## 📌 Что делать, если пользователь просит расширения

| Запрос | Что делает ассистент |
|---|---|
| **Режим SUMIF** (устойчивость к сортировке) | Добавляет вспомогательный столбец «путь группы» и заменяет `SUM` на `SUMIF`; описывает алгоритм. |
| **Асинхронные задачи** (много файлов) | Добавляет Celery/RQ + Redis; `/convert` возвращает `job_id`, `/result/{id}` отдаёт файл. |
| **HTTPS** | Добавляет nginx-конфиг перед контейнером + сертификаты. |
| **Rate-limit** | Подключает slowapi, ставит лимит на IP/ключ. |
| **Несколько листов** | Добавляет параметр `all_sheets=true` и цикл по `wb.sheetnames`. |
| **CI/CD** | Пишет `.github/workflows/build.yml` со сборкой образа и пушем в registry. |

---

Готово. Такой пакет можно целиком отдать ассистенту (Claude, GPT, локальной модели) — он выдаст рабочий проект без уточняющих вопросов. Если нужно — сделаю такую же упаковку для режима **SUMIF** или для **асинхронной версии** с очередью.
