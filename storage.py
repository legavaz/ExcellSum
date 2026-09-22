import asyncio
import os
import shutil
import time
import uuid
from pathlib import Path

STORAGE_DIR = Path(os.getenv("STORAGE_DIR", "/data/storage"))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)
TTL_SECONDS = int(os.getenv("STORAGE_TTL_HOURS", "24")) * 3600
CLEANUP_INTERVAL_SECONDS = 600


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


def purge_expired(now: float | None = None) -> int:
    now = time.time() if now is None else now
    removed = 0
    for d in STORAGE_DIR.iterdir():
        try:
            if d.is_dir() and now - d.stat().st_mtime > TTL_SECONDS:
                shutil.rmtree(d, ignore_errors=True)
                removed += 1
        except OSError:
            pass
    return removed


async def cleanup_loop():
    while True:
        purge_expired()
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)
