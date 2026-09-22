import json
import sys
from pathlib import Path

from app import app

out = Path(sys.argv[1] if len(sys.argv) > 1 else "openapi.json")
out.write_text(json.dumps(app.openapi(), ensure_ascii=False, indent=2),
               encoding="utf-8")
print(f"OpenAPI schema -> {out.resolve()}")
