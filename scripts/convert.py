#!/usr/bin/env python3
"""
Converts the weekly arrears xlsx files into data.json + data.js for the portal.

Looks in the data/ folder for the newest:
  - "Del Report*.xlsx"      -> FDGL / Paytek tabs (split by PSAVE in Lease No.)
                               PS Team & New in Arrears sheets become row tags
  - "PS Lease*.xlsx"        -> PS Lease tab

Settled tracking: before overwriting data.json, the previous version is read.
Any lease that was present last time but is missing from the new report is
moved to the cumulative "Settled" tab (and removed again if it ever reappears).

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

SETTLED_COLS = ["Provider", "MID", "Legal Name", "Trading Name", "Lease No.",
                "Last Arrears", "Paid", "Last Status", "Settled On"]


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


def col(columns, *names):
    low = [c.lower() for c in columns]
    for n in names:
        if n in low:
            return low.index(n)
    return -1


def lease_keys(tab):
    """set of lease numbers in a tab dict"""
    i = col(tab["columns"], "lease no.", "lease")
    if i < 0:
        return set()
    return {str(r[i]) for r in tab["rows"] if i < len(r)}


def settled_entry(provider, columns, row, settled_on):
    g = lambda *names: (row[col(columns, *names)]
                        if col(columns, *names) > -1 and col(columns, *names) < len(row) else "")
    return [provider,
            g("mid"), g("legal name"), g("trading name"),
            g("lease no.", "lease"), g("arrears"),
            g("paid", "payments made"),
            g("lease status", "mandate status", "back office stage"),
            settled_on]


def main():
    del_file = newest(r"del report")
    ps_file = newest(r"ps lease")

    if not del_file and not ps_file:
        sys.exit("No xlsx files found in data/ - nothing to do.")

    # previous state (for settled tracking)
    prev = None
    if OUT_FILE.exists():
        try:
            prev = json.loads(OUT_FILE.read_text(encoding="utf-8"))
        except Exception:
            prev = None

    tabs = {}
    sources = {}

    if del_file:
        print(f"Del Report file : {del_file.name}")
        sources["del_report"] = {"file": del_file.name,
                                 "date": file_date(del_file).strftime("%d/%m/%Y")}
        wb = openpyxl.load_workbook(del_file, read_only=True, data_only=True)

        main_ws = wb.worksheets[0]
        header, rows = read_sheet(main_ws)
        lease_idx = col(header, "lease no.")
        if lease_idx < 0:
            lease_idx = 3

        def sheet_leases(name):
            if name not in wb.sheetnames:
                return set()
            th, tr = read_sheet(wb[name])
            tl = col(th, "lease no.")
            if tl < 0:
                tl = 3
            return {str(r[tl]) for r in tr if tl < len(r)}

        ps_team_leases = sheet_leases("PS Team")
        new_arr_leases = sheet_leases("New in Arrears")

        header = header + ["PS Team", "New in Arrears"]
        fdgl, paytek = [], []
        for r in rows:
            lease_raw = str(r[lease_idx]) if lease_idx < len(r) else ""
            r = list(r) + ["Yes" if lease_raw in ps_team_leases else "",
                           "Yes" if lease_raw in new_arr_leases else ""]
            (paytek if "PSAVE" in lease_raw.upper() else fdgl).append(r)
        tabs["fdgl"] = {"label": "FDGL", "columns": header, "rows": fdgl}
        tabs["paytek"] = {"label": "Paytek", "columns": header, "rows": paytek}
        wb.close()

    if ps_file:
        print(f"PS Lease file   : {ps_file.name}")
        sources["ps_lease"] = {"file": ps_file.name,
                               "date": file_date(ps_file).strftime("%d/%m/%Y")}
        wb = openpyxl.load_workbook(ps_file, read_only=True, data_only=True)
        h, r = read_sheet(wb.worksheets[0])
        tabs["ps_lease"] = {"label": "PS Lease", "columns": h, "rows": r}
        wb.close()

    # ---------- settled tracking ----------
    PROVIDERS = ["fdgl", "paytek", "ps_lease"]
    all_new_keys = set()
    for k in PROVIDERS:
        if k in tabs:
            all_new_keys |= lease_keys(tabs[k])

    settled_rows = []
    seen = set()

    # carry forward previously settled leases (unless they reappeared)
    if prev and "settled" in prev.get("tabs", {}):
        li = SETTLED_COLS.index("Lease No.")
        for r in prev["tabs"]["settled"]["rows"]:
            key = str(r[li]) if li < len(r) else ""
            if key and key not in all_new_keys and key not in seen:
                settled_rows.append(r)
                seen.add(key)

    # newly settled = in previous report but not in the new one
    if prev:
        for k in PROVIDERS:
            pt = prev.get("tabs", {}).get(k)
            nt = tabs.get(k)
            if not pt or not nt:
                continue
            settled_on = sources.get("del_report" if k != "ps_lease" else "ps_lease",
                                     {}).get("date", datetime.now().strftime("%d/%m/%Y"))
            new_keys = lease_keys(nt)
            li = col(pt["columns"], "lease no.", "lease")
            if li < 0:
                continue
            for r in pt["rows"]:
                key = str(r[li]) if li < len(r) else ""
                if key and key not in new_keys and key not in seen:
                    settled_rows.append(settled_entry(pt["label"], pt["columns"], r, settled_on))
                    seen.add(key)

    tabs["settled"] = {"label": "Settled", "columns": SETTLED_COLS, "rows": settled_rows}

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
