#!/usr/bin/env python3
"""
Converts the weekly arrears xlsx files into data.json + data.js for the portal.

Looks in the data/ folder for the newest:
  - "Del Report*.xlsx"      -> FDGL / Paytek / PS Team / New in Arrears tabs
  - "PS Lease*.xlsx"        -> PS Lease tab

Split rule: Lease No. containing "PSAVE" = Paytek, otherwise FDGL.

Run:  python scripts/convert.py
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
        sys.exit("No xlsx files found in data/ - nothing to do.")

    tabs = {}
    sources = {}

    if del_file:
        print(f"Del Report file : {del_file.name}")
        sources["del_report"] = {"file": del_file.name,
                                 "date": file_date(del_file).strftime("%d/%m/%Y")}
        wb = openpyxl.load_workbook(del_file, read_only=True, data_only=True)

        main_ws = wb.worksheets[0]
        header, rows = read_sheet(main_ws)
        try:
            lease_idx = [h.lower() for h in header].index("lease no.")
        except ValueError:
            lease_idx = 3
        fdgl, paytek = [], []
        for r in rows:
            lease = str(r[lease_idx]).upper() if lease_idx < len(r) else ""
            (paytek if "PSAVE" in lease else fdgl).append(r)
        tabs["fdgl"] = {"label": "FDGL", "columns": header, "rows": fdgl}
        tabs["paytek"] = {"label": "Paytek", "columns": header, "rows": paytek}

        for sheet_name, key, label in [("PS Team", "ps_team", "PS Team"),
                                       ("New in Arrears", "new_arrears", "New in Arrears")]:
            if sheet_name in wb.sheetnames:
                h, r = read_sheet(wb[sheet_name])
                tabs[key] = {"label": label, "columns": h, "rows": r}
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
    (ROOT / "data.js").write_text("window.PORTAL_DATA=" + payload + ";",
                                  encoding="utf-8")

    for k, t in tabs.items():
        print(f"  {t['label']:<15} {len(t['rows'])} rows")
    print(f"Wrote {OUT_FILE.name} and data.js")


if __name__ == "__main__":
    main()
