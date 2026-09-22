import os
import tempfile
from pathlib import Path

os.environ["STORAGE_DIR"] = str(
    Path(tempfile.gettempdir()) / "subtotals-api-test-storage"
)
os.environ["STORAGE_TTL_HOURS"] = "1"
