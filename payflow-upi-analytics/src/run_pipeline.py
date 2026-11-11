"""Rebuild everything from scratch: data -> Excel workbook -> (optional) formula check -> web dashboard.

Run:  python -m src.run_pipeline            (about 1 minute)
      python -m src.run_pipeline --verify   (also recalculates the workbook in LibreOffice and compares every formula)
"""
from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

from . import build_excel, build_web, generate_data

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    t0 = time.time()
    generate_data.main()
    build_excel.build()
    page = build_web.build()
    docs = ROOT / "docs"
    docs.mkdir(exist_ok=True)
    shutil.copy(page, docs / "index.html")          # GitHub Pages can serve the /docs folder
    status = 0
    if "--verify" in sys.argv:
        from . import verify_excel
        status = verify_excel.main()
    print(f"Done in {time.time() - t0:.0f}s")
    return status


if __name__ == "__main__":
    sys.exit(main())
