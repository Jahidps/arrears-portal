#!/usr/bin/env python3
"""
Converts the weekly arrears xlsx files into data.json for the portal.

Looks in the data/ folder for the newest:
  - "Del Report*.xlsx"      -> FDGL / Paytek / PS Team / New in Arrears tabs
  - "PS Lease*.xlsx"        -> PS Lease tab

Split rule: Lease No. containing "PSAVE" = Paytek, otherwise FDGL.

Run:  python scripts/convert.py
Output: data.json in repo root.
"""
import json
import re
import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
OUT_FILE = ROOT / "data.json"

DATE_RE = re.compile(r"(\d{2})-(\d{2})-(\d{4})")


def file_date(path: Path):
    m = DATE_RE.search(path.name)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return datetime(y, mo, d)
        except ValueError:
            pass
    return datetime.fromtimestamp(path.stat().st_mtime)


def newest(pattern: str):
    files = [p for p in DATA_DIR.glob("*.xlsx")
             if re.match(pattern, p.name, re.IGNORECASE) and not p.name.startswith("~")]
    if not files:
        return None
    return max(files, key=file_date)


def clean(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v


def read_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    # drop fully-empty trailing columns
    while header and header[-1] == "":
        header.pop()
    data = []
    for r in rows[1:]:
        vals = [clean(v) for v in r[: len(header)]]
        if any(v != "" for v in vals):
            data.append(vals)
    return header, data


def main():
    del_file = newest(r"del report")
    ps_file = newest(r"ps lease")

    if not del_file and not ps_file:
        sys.exit("No xlsx files found in data/ — nothing to do.")

    tabs = {}
    sources = {}

    if del_file:
        print(f"Del Report file : {del_file.name}")
        sources["del_report"] = {"file": del_file.name,
                                 "date": file_date(del_file).strftime("%d/%m/%Y")}
        wb = openpyxl.load_workbook(del_file, read_only=True, data_only=True)

        # Main sheet = first sheet
        main_ws = wb.worksheets[0]
        header, rows = read_sheet(main_ws)
        try:
            lease_idx = [h.lower() for h in header].index("lease no.")
        except ValueError:
            lease_idx = 3

        # PS Team sheet rows are a subset of the main sheet -> tag them
        ps_team_leases = set()
        if "PS Team" in wb.sheetnames:
            th, tr = read_sheet(wb["PS Team"])
            try:
                tl = [h.lower() for h in th].index("lease no.")
            except ValueError:
                tl = 3
            ps_team_leases = {str(r[tl]) for r in tr if tl < len(r)}

        header = header + ["PS Team"]
        fdgl, paytek = [], []
        for r in rows:
            lease = str(r[lease_idx]).upper() if lease_idx < len(r) else ""
            r = list(r) + ["Yes" if str(r[lease_idx]) in ps_team_leases else ""]
            (paytek if "PSAVE" in lease else fdgl).append(r)
        tabs["fdgl"] = {"label": "FDGL", "columns": header, "rows": fdgl}
        tabs["paytek"] = {"label": "Paytek", "columns": header, "rows": paytek}

        if "New in Arrears" in wb.sheetnames:
            h, r = read_sheet(wb["New in Arrears"])
            tabs["new_arrears"] = {"label": "New in Arrears", "columns": h, "rows": r}
        wb.close()

    if ps_file:
        print(f"PS Lease file   : {ps_file.name}")
        sources["ps_lease"] = {"file": ps_file.name,
                               "date": file_date(ps_file).strftime("%d/%m/%Y")}
        wb = openpyxl.load_workbook(ps_file, read_only=True, data_only=True)
        h, r = read_sheet(wb.worksheets[0])
        tabs["ps_lease"] = {"label": "PS Lease", "columns": h, "rows": r}
        wb.close()

    out = {
        "generated": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "sources": sources,
        "tabs": tabs,
    }
    payload = json.dumps(out, ensure_ascii=False, separators=(",", ":"))
    OUT_FILE.write_text(payload, encoding="utf-8")
    # data.js lets the portal work when opened as a local file (no fetch/CORS)
    (ROOT / "data.js").write_text("window.PORTAL_DATA=" + payload + ";",
                                  encoding="utf-8")
  
    for k, t in tabs.items():
        print(f"  {t['label']:<15} {len(t['rows'])} rows")
    print(f"Wrote {OUT_FILE.name} and data.js")


if __name__ == "__main__":
    main()
