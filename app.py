from contextlib import asynccontextmanager
from pathlib import Path
import asyncio
import base64
import contextlib
import io

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl.utils import column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException

from openapi_meta import API_DESCRIPTION, CONTACT, LICENSE_INFO, tags_metadata
from schemas import AnalyzeResponse, ErrorResponse, HealthResponse
from storage import cleanup_loop, new_job_dir, save_upload
from subtotals import SheetNotFound, analyze, parse_columns, process

WEB_DIR = Path(__file__).resolve().parent / "web"


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(cleanup_loop())
    try:
        yield
    finally:
        task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await task


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
    expose_headers=["X-Log-Preview", "X-Log-Preview-B64", "X-Job-Id"],
)


def _parse_level_column(s):
    if not s:
        return None
    s = s.strip()
    return int(s) if s.isdigit() else column_index_from_string(s)


def _validate_style(s):
    if s not in ("auto", "top", "bottom"):
        raise HTTPException(400, "style must be auto|top|bottom")
    return None if s == "auto" else s


def _header_safe(value, limit=1500):
    text = value[-limit:].replace("\r", " ").replace("\n", " | ")
    return text.encode("latin-1", "replace").decode("latin-1")


def _log_headers(log):
    tail = log[-1500:]
    return {
        "X-Log-Preview": _header_safe(log),
        "X-Log-Preview-B64": base64.b64encode(tail.encode("utf-8")).decode("ascii"),
    }


def _run_conversion(src, dst, columns, level_col, style, sheet, dry_run,
                    grand_total):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        process(str(src), str(dst), value_columns=columns,
                level_column=level_col, sheet_name=sheet,
                forced_style=style, dry_run=dry_run,
                grand_total=grand_total)
    return buf.getvalue()


@app.get("/", include_in_schema=False, response_class=FileResponse,
         summary="Веб-форма конвертации")
def ui():
    return FileResponse(WEB_DIR / "index.html", media_type="text/html")


@app.get("/health", tags=["system"], response_model=HealthResponse,
         summary="Проверка живости сервиса")
def health():
    return {"status": "ok"}


@app.post("/analyze", tags=["convert"], response_model=AnalyzeResponse,
          summary="Разобрать файл и подсказать настройки",
          responses={400: {"model": ErrorResponse},
                     500: {"model": ErrorResponse}})
async def analyze_file(
    file: UploadFile = File(..., description=".xlsx из 1С"),
    sheet: str | None = Form(None),
    level_column: str | None = Form(None),
):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx files are supported")

    content = await file.read()
    try:
        level_col = _parse_level_column(level_column)
        return analyze(io.BytesIO(content), sheet, level_col)
    except SheetNotFound as e:
        raise HTTPException(400, str(e))
    except InvalidFileException:
        raise HTTPException(400, "Файл повреждён или не является .xlsx")
    except Exception as e:
        raise HTTPException(500, f"Analyze failed: {e}")


@app.post("/convert", tags=["convert"],
          summary="Заменить итоги на формулы Excel",
          response_class=FileResponse,
          responses={
              200: {"description": "Готовый .xlsx либо JSON при dry_run=true",
                    "content": {
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                            {"schema": {"type": "string", "format": "binary"}},
                        "application/json":
                            {"schema": {"$ref": "#/components/schemas/DryRunResponse"}},
                    }},
              400: {"model": ErrorResponse},
              401: {"model": ErrorResponse},
              500: {"model": ErrorResponse},
          })
async def convert(
    file: UploadFile = File(..., description=".xlsx из 1С"),
    columns: str | None = Form(
        None, description="Колонки со значениями, напр. E,F,G; пусто — автоопределение"),
    level_column: str | None = Form(None),
    style: str = Form("auto"),
    sheet: str | None = Form(None),
    dry_run: bool = Form(False),
    grand_total: bool = Form(False),
):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx files are supported")

    try:
        cols = parse_columns(columns or "")
        level_col = _parse_level_column(level_column)
        forced_style = _validate_style(style)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"Bad parameters: {e}")

    job_dir = new_job_dir()
    src = save_upload(job_dir, file.filename, file.file)
    dst = job_dir / f"fixed_{src.name}"

    try:
        log = _run_conversion(src, dst, cols, level_col, forced_style, sheet,
                              dry_run=dry_run, grand_total=grand_total)
    except SheetNotFound as e:
        raise HTTPException(400, str(e))
    except InvalidFileException:
        raise HTTPException(400, "Файл повреждён или не является .xlsx")
    except Exception as e:
        raise HTTPException(500, f"Conversion failed: {e}")

    if dry_run:
        return JSONResponse(
            {"log": log, "dry_run": True},
            headers={"X-Job-Id": job_dir.name, **_log_headers(log)},
        )

    if not dst.exists():
        raise HTTPException(500, "Output file was not created")

    return FileResponse(
        dst, filename=f"fixed_{file.filename}",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            **_log_headers(log),
            "X-Job-Id": job_dir.name,
        },
    )
