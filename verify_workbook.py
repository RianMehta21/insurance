#!/usr/bin/env python3
"""
Standalone workbook verifier.

Forces LibreOffice to recalculate every formula in an .xlsx, then scans the
result for error cells. Exits non-zero if any are found, so CI can refuse to
publish a broken workbook.

Why this exists: openpyxl writes formula STRINGS. It never evaluates them. A
workbook can be saved with 7,500 formulas that are all silently broken and
openpyxl will report nothing wrong. The only way to know they work is to make a
real spreadsheet engine calculate them.

Requires LibreOffice:
    Ubuntu/Debian   sudo apt-get install -y libreoffice-calc
    macOS           brew install --cask libreoffice
    Windows         install LibreOffice, then add the program folder to PATH

Usage:
    python verify_workbook.py VHIS_Compare.xlsx
    python verify_workbook.py VHIS_Compare.xlsx --timeout 300 --json
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

# The seven values Excel and LibreOffice use to signal a broken formula.
ERROR_VALUES = {
    "#REF!":   "deleted or out-of-range reference",
    "#VALUE!": "wrong data type in an argument",
    "#DIV/0!": "division by zero",
    "#N/A":    "lookup found no match",
    "#NAME?":  "unrecognised function or defined name",
    "#NULL!":  "empty intersection of two ranges",
    "#NUM!":   "invalid number for the operation",
}

# StarBasic. calculateAll() forces a full recalculation regardless of the
# document's AutoCalculate setting, which a plain --convert-to does not
# guarantee. This is the whole reason we inject a macro instead of just
# round-tripping the file through soffice.
MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script"
               script:name="Module1" script:language="StarBasic">
Sub RecalculateAndSave
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>
"""

def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def install_macro(soffice, profile: Path, timeout=90):
    """Two phases, and the order is not optional.

    LibreOffice must build its OWN profile first (--terminate_after_init).
    A hand-built directory tree looks superficially right but is missing the
    internal registration LibreOffice expects, and the macro then never
    resolves: soffice launches, finds nothing to run, and hangs until killed.
    Only once the real profile exists do we drop the module into it.
    """
    url = profile.as_uri()
    subprocess.run(
        [soffice, "--headless", "--terminate_after_init",
         f"-env:UserInstallation={url}"],
        capture_output=True, timeout=timeout)

    std = profile / "user" / "basic" / "Standard"
    if not std.exists():
        raise RuntimeError(
            "LibreOffice did not create a usable profile. Formulas were NOT "
            "verified. Check that libreoffice-calc is properly installed.")
    (std / "Module1.xba").write_text(MACRO, encoding="utf-8")
    return url


def recalculate(path: Path, timeout: int):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found on PATH. Install libreoffice-calc, or skip "
            "verification with --no-recalc (not recommended for CI).")

    with tempfile.TemporaryDirectory(prefix="verify-lo-") as tmp:
        profile = Path(tmp) / "profile"
        url = install_macro(soffice, profile)

        # Recalculation rewrites in place, so work on a copy and leave the
        # user's file untouched.
        work = Path(tmp) / path.name
        shutil.copy2(path, work)

        cmd = [soffice, "--headless", "--norestore",
               f"-env:UserInstallation={url}",
               "vnd.sun.star.script:Standard.Module1.RecalculateAndSave"
               "?language=Basic&location=application",
               str(work.absolute())]
        if shutil.which("timeout"):
            cmd = ["timeout", str(timeout)] + cmd

        before = work.stat().st_mtime_ns
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout + 15)
        if work.stat().st_mtime_ns == before:
            raise RuntimeError(
                "LibreOffice did not rewrite the file, so nothing was "
                "recalculated. stderr: "
                + proc.stderr.decode(errors="replace")[:300])

        out = path.with_suffix(".recalc.xlsx")
        shutil.copy2(work, out)
        return out


def scan(path: Path, max_locations=100):
    wb = load_workbook(path, data_only=True)
    counts, locations = {}, []
    total_cells = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                total_cells += 1
                if isinstance(v, str) and v in ERROR_VALUES:
                    counts[v] = counts.get(v, 0) + 1
                    if len(locations) < max_locations:
                        locations.append(f"{ws.title}!{cell.coordinate} {v}")
    wb.close()
    return counts, locations, total_cells


def count_formulas(path: Path):
    wb = load_workbook(path, data_only=False)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    n += 1
    wb.close()
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--timeout", type=int, default=300)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--keep", action="store_true",
                    help="keep the recalculated copy for inspection")
    args = ap.parse_args()

    path = Path(args.workbook)
    if not path.exists():
        print(f"not found: {path}")
        sys.exit(2)

    formulas = count_formulas(path)
    recalced = recalculate(path, args.timeout)
    counts, locations, cells = scan(recalced)
    if not args.keep:
        recalced.unlink(missing_ok=True)

    total = sum(counts.values())
    result = {
        "status": "success" if total == 0 else "errors_found",
        "workbook": str(path),
        "total_formulas": formulas,
        "populated_cells": cells,
        "total_errors": total,
        "error_summary": counts,
        "locations": locations,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"workbook   : {path}")
        print(f"formulas   : {formulas:,}")
        print(f"cells      : {cells:,}")
        print(f"errors     : {total}")
        if counts:
            print("\nbreakdown:")
            for k, v in sorted(counts.items(), key=lambda x: -x[1]):
                print(f"  {k:<9} {v:>6}   {ERROR_VALUES[k]}")
            print("\nfirst locations:")
            for loc in locations[:20]:
                print("  ", loc)
        else:
            print("\nAll formulas evaluated cleanly.")

    sys.exit(0 if total == 0 else 1)


if __name__ == "__main__":
    main()
