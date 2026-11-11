"""Verify the workbook: recalculate every formula in LibreOffice and compare with the cached values.

A mismatch means a formula does not compute what the Python analysis says it should, or an Excel
function is unsupported. Also checks for error values (#NAME?, #DIV/0!, #VALUE!, ...).

Run:  python -m src.verify_excel
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "excel" / "PayFlow_UPI_Analytics.xlsx"
SKIP = {"Data"}

RECALC_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def recalc_copy(src: Path, workdir: Path) -> Path:
    profile = workdir / "profile"
    (profile / "user").mkdir(parents=True)
    (profile / "user" / "registrymodifications.xcu").write_text(RECALC_XCU)
    out = workdir / "out"
    out.mkdir()
    cmd = ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
           "--outdir", str(out), str(src)]
    subprocess.run(cmd, check=True, timeout=900, capture_output=True)
    return out / src.name


def same(a, b) -> bool:
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-6 * max(1.0, abs(a), abs(b))
    return str(a) == str(b)


def main() -> int:
    work = Path(tempfile.mkdtemp(prefix="pf_verify_"))
    try:
        print("[verify] recalculating in LibreOffice (this takes a minute)...")
        recalculated = recalc_copy(XLSX, work)
        f_wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=False)
        c_wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
        r_wb = openpyxl.load_workbook(recalculated, read_only=True, data_only=True)
        total = bad = errors = 0
        examples = []
        for name in f_wb.sheetnames:
            if name in SKIP:
                continue
            fr = list(f_wb[name].iter_rows())
            cr = list(c_wb[name].iter_rows(values_only=True))
            rr = list(r_wb[name].iter_rows(values_only=True))
            for i, row in enumerate(fr):
                for j, cell in enumerate(row):
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    total += 1
                    cached = cr[i][j] if i < len(cr) and j < len(cr[i]) else None
                    fresh = rr[i][j] if i < len(rr) and j < len(rr[i]) else None
                    if isinstance(fresh, str) and fresh.startswith("#"):
                        errors += 1
                        examples.append((name, cell.coordinate, v[:70], cached, fresh))
                    elif not same(cached, fresh):
                        bad += 1
                        if len(examples) < 15:
                            examples.append((name, cell.coordinate, v[:70], cached, fresh))
        print(f"[verify] formula cells checked: {total:,} | mismatches: {bad} | error values: {errors}")
        for e in examples[:15]:
            print("   ", e)
        return 0 if (bad == 0 and errors == 0) else 1
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
