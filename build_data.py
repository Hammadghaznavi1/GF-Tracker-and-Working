#!/usr/bin/env python3
"""Build the dashboard data from the Excel workbook and prepare the website folder.

Run by GitHub Actions on every push (so uploading a new Excel file updates the live site).
Local test:  DASHBOARD_PASSWORD=secret python scripts/build_data.py
Output:      site/index.html, site/app.js, site/data.json (encrypted with the password)
"""
import argparse
import base64
import datetime as dt
import glob
import hashlib
import json
import os
import re
import shutil
import sys
from pathlib import Path

import openpyxl

PG = ["Financial Inclusion Programme", "Ehsaas", "Cash Transfer Programme", "G2P Digitisation Programme",
      "Cash in Cash Out (CICO)", "Raast Regional Connectivity", "FBR Digitalisation"]
CODE = ["INV-007619", "INV-032192", "INV-061756", "INV-087029", "INV-087089", "INV-086432", "INV-102619"]
PROG_MAP = {"financial inclusion program": 0, "ehsaas program": 1, "cash transfer program": 2,
            "g2p digitization program": 3, "cash in cash out (cico)": 4,
            "raast regional connectivity": 5, "fbr digitalization": 6}
DETAIL_COLS = {"E": 0, "F": 0, "G": 1, "H": 2, "I": 3, "J": 4, "K": 5, "L": 6}   # F = Mckinsey-tagged
PROFIT_COLS = {0: [("C", "D"), ("E", "F")], 1: [("G", "H")], 2: [("I", "J")], 3: [("K", "L")],
               4: [("M", "N")], 5: [("O", "P")], 6: [("Q", "R")]}
DAYS = [31, 31, 30]          # July, August, September (profit on actual/360)
ITER = 250_000               # PBKDF2 iterations


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%d-%m-%y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    sys.exit(f"ERROR: sheet '{name}' not found. Sheets in the file: {[w.title for w in wb.worksheets]}")


def find_workbook(arg):
    if arg:
        return arg
    if Path("data/GF_Balance.xlsx").exists():
        return "data/GF_Balance.xlsx"
    files = glob.glob("data/*.xlsx")
    if len(files) == 1:
        return files[0]
    sys.exit("ERROR: put exactly one Excel file in the data/ folder (ideally named GF_Balance.xlsx). "
             f"Found: {files}")


def build(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    # ---- receipts ----
    GF, FC = [], []
    for r in sheet(wb, "EAGR & GF").iter_rows(values_only=True):
        if len(r) < 11:
            continue
        donor, d = str(r[2] or "").strip(), parse_date(r[3])
        if not d:
            continue
        if donor == "GF" and num(r[6]):
            i = PROG_MAP.get(str(r[4]).strip().lower())
            if i is not None:
                GF.append([d.isoformat(), i, num(r[6]), round(num(r[8]), 4), round(num(r[9]), 2)])
        elif donor == "FCDO":
            FC.append([d.isoformat(), str(r[4]).strip(), str(r[5] or "").strip(), round(num(r[7]), 2),
                       round(num(r[8]), 4), round(num(r[10]), 2)])
    GF.sort()
    FC.sort()
    if not GF:
        sys.exit("ERROR: no GF receipts found in sheet 'EAGR & GF'.")

    # ---- monthly balance blocks ----
    ws = sheet(wb, "26-27 Detailed")

    def row7(r):
        out = [0.0] * 7
        for col, i in DETAIL_COLS.items():
            out[i] += num(ws[f"{col}{r}"].value)
        return out

    blocks, cur = [], None
    for r in range(1, ws.max_row + 1):
        lab = str(ws.cell(r, 3).value or "").strip()
        if lab.startswith("Opening bank balances in PKR"):
            cur = {"conv": [0.0] * 7, "eq": [0.0] * 7}
            blocks.append(cur)
        if cur is None or not lab:
            continue
        if lab.startswith("Opening bank balances in PKR"):
            cur["pkr_open"] = row7(r)
        elif lab.startswith("Opening bank balances in USD"):
            cur["usd_open"] = row7(r)
        elif lab.startswith("Opening short term investment TDR"):
            cur["tdr_open"] = row7(r)
        elif lab.startswith("Add: Equity Injection"):
            cur["eq"] = row7(r)
        elif lab.startswith("Less: Outflow from USD bank account"):
            cur["conv"] = [a - b for a, b in zip(cur["conv"], row7(r))]
        elif lab.startswith("Closing bank balances in PKR"):
            cur["pkr_close"] = row7(r)
        elif lab.startswith("Closing bank balances in USD"):
            cur["usd_close"] = row7(r)
            cur["as_of"] = lab.split(" as of ")[-1]
        elif lab.startswith("Closing short term investment TDR"):
            cur["tdr_close"] = row7(r)
    blocks = blocks[:4]
    if len(blocks) < 4:
        sys.exit(f"ERROR: expected 4 monthly blocks (Jul, Aug, Sep, Oct) in '26-27 Detailed', found {len(blocks)}.")
    for b in blocks:
        for k in ("pkr_open", "usd_open", "tdr_open", "pkr_close", "usd_close", "tdr_close"):
            b.setdefault(k, [0.0] * 7)
    S = lambda b, k: sum(b[k])

    # ---- profit split ----
    ps = sheet(wb, "GF Profit - USD PKR Split 26-27")
    rate = [num(ps[f"U{r}"].value) for r in (5, 6, 7)]
    pu = [[sum(num(ps[f"{a}{r}"].value) for a, _ in PROFIT_COLS[i]) for r in (5, 6, 7)] for i in range(7)]
    pk = [[sum(num(ps[f"{b}{r}"].value) for _, b in PROFIT_COLS[i]) for r in (5, 6, 7)] for i in range(7)]
    profit = sum(map(sum, pu)) + sum(sum(pk[i][m] / rate[m] for i in range(7)) for m in range(3))

    openA = [S(blocks[m], "usd_open") + S(blocks[m], "tdr_open") for m in range(3)]
    closeA = [S(blocks[m], "usd_close") + S(blocks[m], "tdr_close") for m in range(3)]
    avgbal = sum(openA[m] * DAYS[m] for m in range(3)) / 92
    monthly = [list(b["conv"]) for b in blocks]

    # ---- dates for captions ----
    as_of, lab = None, blocks[3].get("as_of", "")
    m = re.search(r"([A-Za-z]+)\s+(\d+)(?:st|nd|rd|th)?,?\s+(\d{4})", lab)     # "October 5th, 2026"
    m2 = re.search(r"(\d+)(?:st|nd|rd|th)?\s+([A-Za-z]+),?\s+(\d{4})", lab)    # "5th October 2026"
    for mon, day, yr in ([(m.group(1), m.group(2), m.group(3))] if m else []) + ([(m2.group(2), m2.group(1), m2.group(3))] if m2 else []):
        try:
            as_of = dt.datetime.strptime(f"{day} {mon} {yr}", "%d %B %Y").date()
            break
        except ValueError:
            pass
    last_gf = dt.date.fromisoformat(GF[-1][0])
    long = lambda d: f"{d.day} {d.strftime('%B %Y')}"
    short = lambda d: f"{d.day} {d.strftime('%b %Y')}"
    as_of = as_of or last_gf
    oct_dates = [dt.date.fromisoformat(g[0]) for g in GF
                 if dt.date.fromisoformat(g[0]).year == as_of.year and dt.date.fromisoformat(g[0]).month == as_of.month]
    F = {
        "PG": PG, "CODE": CODE,
        "open": [blocks[0]["usd_open"][i] + blocks[0]["tdr_open"][i] for i in range(7)],
        "avail": list(blocks[3]["usd_close"]), "tdr": list(blocks[3]["tdr_close"]),
        "pkr": list(blocks[2]["pkr_close"]),
        "spend": [sum(b["conv"][i] for b in blocks) for i in range(7)],
        "monthly": monthly, "pu": pu, "pk": pk, "rate": rate,
        "yld": profit / avgbal * 360 / 92, "latest": GF[-1][3], "avgbal": avgbal, "q1days": 92,
        "openA": openA, "closeA": closeA, "convA": [sum(monthly[m]) for m in range(3)],
        "eqA": [S(blocks[m], "eq") for m in range(3)],
        "octEq": S(blocks[3], "eq"), "octDate": long(oct_dates[-1]) if oct_dates else "",
        "asOf": long(as_of), "asOfShort": short(as_of), "lastGF": long(last_gf),
        "built": dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d%H%M%S"),
    }
    return {"GF": GF, "FC": FC, "F": F}


def encrypt(obj, password):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv = os.urandom(16), os.urandom(12)
    key = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, ITER, 32)
    ct = AESGCM(key).encrypt(iv, json.dumps(obj, separators=(",", ":")).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"enc": 1, "iter": ITER, "salt": b(salt), "iv": b(iv), "ct": b(ct)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx")
    ap.add_argument("--site", default="site")
    a = ap.parse_args()
    path = find_workbook(a.xlsx)
    print(f"Reading {path}")
    data = build(path)
    pw = os.environ.get("DASHBOARD_PASSWORD", "")
    if pw:
        out = encrypt(data, pw)
        print("Data encrypted with DASHBOARD_PASSWORD.")
    elif os.environ.get("ALLOW_PLAIN") == "1":
        out = data
        print("WARNING: data is NOT encrypted (ALLOW_PLAIN=1).")
    else:
        sys.exit("ERROR: DASHBOARD_PASSWORD is not set. Add it as a GitHub secret so the published data is encrypted.")
    site = Path(a.site)
    site.mkdir(exist_ok=True)
    for f in ("index.html", "app.js"):
        shutil.copy(Path("web") / f, site / f)
    (site / ".nojekyll").write_text("")
    (site / "data.json").write_text(json.dumps(out, separators=(",", ":")))
    print(f"Built {site}/ : {len(data['GF'])} GF receipts, {len(data['FC'])} FCDO receipts, "
          f"as of {data['F']['asOf']}")


if __name__ == "__main__":
    main()
