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
import tempfile
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
DAYS12 = [31, 31, 30, 31, 30, 31, 31, 28, 31, 30, 31, 30]   # Jul to Jun (profit on actual/360)
ITER = 250_000               # PBKDF2 iterations


def num(v):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").replace("$", "").replace("£", "").strip())
        except ValueError:
            return 0.0
    return 0.0


DATE_FORMATS = ("%d-%m-%y", "%d-%m-%Y", "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d %B %Y", "%d %b %Y",
                "%d/%m/%Y", "%d/%m/%y", "%d.%m.%Y", "%B %d %Y", "%b %d %Y")


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 30000 < v < 80000:   # Excel serial number
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    if isinstance(v, str):
        t = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", v.strip().replace(",", ""))
        for f in DATE_FORMATS:
            try:
                return dt.datetime.strptime(t, f).date()
            except ValueError:
                pass
    return None


def label_date(lab):
    """Read the date from labels such as 'October 31st, 2026' or '31st October 2026'."""
    m = re.search(r"([A-Za-z]+)\s+(\d+)(?:st|nd|rd|th)?,?\s+(\d{4})", lab or "")
    m2 = re.search(r"(\d+)(?:st|nd|rd|th)?\s+([A-Za-z]+),?\s+(\d{4})", lab or "")
    for mon, day, yr in ([(m.group(1), m.group(2), m.group(3))] if m else []) + ([(m2.group(2), m2.group(1), m2.group(3))] if m2 else []):
        try:
            return dt.datetime.strptime(f"{day} {mon} {yr}", "%d %B %Y").date()
        except ValueError:
            pass
    return None


def programme_index(name):
    """Match programme names loosely (spelling, 'Program'/'Programme', stray hyphens)."""
    n = re.sub(r"[^a-z0-9 ]", " ", str(name or "").lower())
    n = " ".join(n.split())
    for key, i in (("financial inclusion", 0), ("ehsaas", 1), ("cash transfer", 2), ("g2p", 3),
                   ("cico", 4), ("cash in cash out", 4), ("raast", 5), ("fbr", 6)):
        if key in n:
            return i
    return None


def donor_of(v):
    n = str(v or "").strip().lower()
    if n in ("gf", "bmgf", "gates", "gates foundation", "bill & melinda gates foundation"):
        return "GF"
    if n.startswith("fcdo") or n == "eagr":
        return "FCDO"
    return ""


def sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    sys.exit(f"ERROR: sheet '{name}' not found. Sheets in the file: {[w.title for w in wb.worksheets]}")


def decrypt_file(path, password):
    """Decrypt GF_Balance.enc (made by encrypt.html) and return the Excel bytes."""
    from cryptography.exceptions import InvalidTag
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    o = json.loads(Path(path).read_text())
    key = hashlib.pbkdf2_hmac("sha256", password.encode(), base64.b64decode(o["salt"]), int(o["iter"]), 32)
    try:
        return AESGCM(key).decrypt(base64.b64decode(o["iv"]), base64.b64decode(o["ct"]), None)
    except InvalidTag:
        sys.exit("ERROR: could not decrypt data/GF_Balance.enc. The password typed in encrypt.html must be exactly "
                 "the same as the GitHub secret DASHBOARD_PASSWORD.")


def find_workbook(arg):
    if arg:
        return arg
    if Path("data/GF_Balance.enc").exists():
        pw = os.environ.get("DASHBOARD_PASSWORD", "")
        if not pw:
            sys.exit("ERROR: data/GF_Balance.enc needs the DASHBOARD_PASSWORD secret to be read.")
        tmp = Path(tempfile.mkdtemp()) / "GF_Balance.xlsx"
        tmp.write_bytes(decrypt_file("data/GF_Balance.enc", pw))
        print("Decrypted data/GF_Balance.enc")
        return str(tmp)
    if Path("data/GF_Balance.xlsx").exists():
        return "data/GF_Balance.xlsx"
    files = glob.glob("data/*.xlsx")
    if len(files) == 1:
        return files[0]
    if len(files) > 1:
        sys.exit(f"ERROR: the data folder has {len(files)} Excel files {files}. Keep exactly one Excel file there.")
    sys.exit("ERROR: no Excel file found. Upload your workbook (.xlsx) into the data folder.")


def build(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    # ---- receipts ----
    GF, FC, skipped = [], [], []
    for n, r in enumerate(sheet(wb, "EAGR & GF").iter_rows(values_only=True), 1):
        if len(r) < 11:
            continue
        donor, d = donor_of(r[2]), parse_date(r[3])
        has_amount = num(r[6]) or num(r[7])
        if not donor:
            if str(r[2] or "").strip() and has_amount and str(r[2]).strip().lower() not in ("donor",):
                skipped.append(f"row {n}: donor '{r[2]}' is not recognised (use GF or FCDO)")
            continue
        if not d:
            if has_amount or str(r[3] or "").strip():
                skipped.append(f"row {n} ({donor}): date '{r[3]}' could not be read (use a real date such as 30-10-26)")
            continue
        if donor == "GF":
            if not num(r[6]):
                skipped.append(f"row {n} (GF {d}): the US$ amount in column G is empty or not a number")
                continue
            i = programme_index(r[4])
            if i is None:
                skipped.append(f"row {n} (GF {d}): programme '{r[4]}' is not recognised")
                continue
            GF.append([d.isoformat(), i, num(r[6]), round(num(r[8]), 4), round(num(r[9]), 2)])
        else:
            FC.append([d.isoformat(), str(r[4]).strip(), str(r[5] or "").strip(), round(num(r[7]), 2),
                       round(num(r[8]), 4), round(num(r[10]), 2)])
    GF.sort()
    FC.sort()
    if not GF:
        sys.exit("ERROR: no GF receipts found in sheet 'EAGR & GF'.")

    # ---- monthly balance blocks (26-27 Detailed) ----
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
            cur = {"conv": [0.0] * 7, "eq": [0.0] * 7, "flow": 0.0}
            blocks.append(cur)
        if cur is None or not lab:
            continue
        low = lab.lower()
        if lab.startswith("Opening bank balances in PKR"):
            cur["pkr_open"] = row7(r)
        elif lab.startswith("Opening bank balances in USD"):
            cur["usd_open"] = row7(r)
        elif lab.startswith("Opening short term investment TDR"):
            cur["tdr_open"] = row7(r)
        elif lab.startswith("Add: Equity Injection"):
            cur["eq"] = row7(r)
            cur["flow"] += sum(abs(x) for x in cur["eq"])
        elif lab.startswith("Less: Outflow from USD bank account"):
            v = row7(r)
            cur["conv"] = [a - b for a, b in zip(cur["conv"], v)]
            cur["flow"] += sum(abs(x) for x in v)
        elif low.startswith("less: outflow from bank to tdr") or low.startswith("add: inflow to bank account due to tdrs") \
                or low.startswith("add: profit earned"):
            cur["flow"] += sum(abs(x) for x in row7(r))
        elif lab.startswith("Closing bank balances in PKR"):
            cur["pkr_close"] = row7(r)
        elif lab.startswith("Closing bank balances in USD"):
            cur["usd_close"] = row7(r)
            cur["as_of"] = lab.split(" as of ")[-1]
        elif lab.startswith("Closing short term investment TDR"):
            cur["tdr_close"] = row7(r)
    blocks = blocks[:12]
    if not blocks:
        sys.exit("ERROR: no monthly blocks ('Opening bank balances in PKR Account...') found in '26-27 Detailed'.")
    for b in blocks:
        for k in ("pkr_open", "usd_open", "tdr_open", "pkr_close", "usd_close", "tdr_close"):
            b.setdefault(k, [0.0] * 7)
        b["date"] = label_date(b.get("as_of", ""))
    first = blocks[0]["date"]
    if not first:
        sys.exit("ERROR: could not read the closing date of the first month in '26-27 Detailed'.")
    y0 = first.year

    def is_month_end(k, b):
        m = (6 + k) % 12 + 1                       # July = block 0
        y = y0 + (6 + k) // 12
        last = (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))
        return b["date"] == last

    nAct = 0                                        # completed months (activity recorded and closed at month end)
    for k, b in enumerate(blocks):
        if b["flow"] > 0 and is_month_end(k, b):
            nAct += 1
        else:
            break
    if nAct < 1:
        sys.exit("ERROR: no completed month found in '26-27 Detailed'.")
    inprog = blocks[nAct] if nAct < len(blocks) and blocks[nAct]["flow"] > 0 else None
    latest = inprog or blocks[nAct - 1]
    S = lambda b, k: sum(b[k])

    # ---- profit split ----
    ps = sheet(wb, "GF Profit - USD PKR Split 26-27")
    rows_ = [5 + m for m in range(nAct)]
    rate = [num(ps[f"U{r}"].value) for r in rows_]
    pu = [[sum(num(ps[f"{a}{r}"].value) for a, _ in PROFIT_COLS[i]) for r in rows_] for i in range(7)]
    pk = [[sum(num(ps[f"{b}{r}"].value) for _, b in PROFIT_COLS[i]) for r in rows_] for i in range(7)]
    if not all(rate):
        sys.exit(f"ERROR: the SBP spot rate (column U) is missing for one of the first {nAct} months in the profit sheet.")
    profit = sum(map(sum, pu)) + sum(sum(pk[i][m] / rate[m] for i in range(7)) for m in range(nAct))

    openA = [S(blocks[m], "usd_open") + S(blocks[m], "tdr_open") for m in range(nAct)]
    closeA = [S(blocks[m], "usd_close") + S(blocks[m], "tdr_close") for m in range(nAct)]
    days = DAYS12[:nAct]
    actDays = sum(days)
    avgbal = sum(openA[m] * days[m] for m in range(nAct)) / actDays
    used = blocks[:nAct + (1 if inprog else 0)]
    monthly = [list(b["conv"]) for b in used]
    pkr_b = latest if sum(abs(x) for x in latest["pkr_close"]) > 0 else blocks[nAct - 1]

    # ---- dates for captions ----
    last_gf = dt.date.fromisoformat(GF[-1][0])
    long = lambda d: f"{d.day} {d.strftime('%B %Y')}"
    short = lambda d: f"{d.day} {d.strftime('%b %Y')}"
    as_of = latest["date"] or last_gf
    cur_dates = []
    if inprog:
        cur_dates = [dt.date.fromisoformat(g[0]) for g in GF
                     if (dt.date.fromisoformat(g[0]).year, dt.date.fromisoformat(g[0]).month) == (as_of.year, as_of.month)
                     and dt.date.fromisoformat(g[0]) <= as_of]
    cur_eq = S(inprog, "eq") if inprog else 0.0
    cur_date = long(cur_dates[-1]) if cur_dates else ""
    F = {
        "PG": PG, "CODE": CODE,
        "nAct": nAct, "actDays": actDays, "q1days": actDays,
        "open": [blocks[0]["usd_open"][i] + blocks[0]["tdr_open"][i] for i in range(7)],
        "avail": list(latest["usd_close"]), "tdr": list(latest["tdr_close"]),
        "pkr": list(pkr_b["pkr_close"]), "pkrAsOf": long(pkr_b["date"]) if pkr_b["date"] else long(as_of),
        "spend": [sum(b["conv"][i] for b in used) for i in range(7)],
        "monthly": monthly, "pu": pu, "pk": pk, "rate": rate,
        "yld": profit / avgbal * 360 / actDays, "latest": GF[-1][3], "avgbal": avgbal,
        "openA": openA, "closeA": closeA, "convA": [sum(m) for m in monthly[:nAct]],
        "eqA": [S(blocks[m], "eq") for m in range(nAct)],
        "curEq": cur_eq, "curDate": cur_date, "octEq": cur_eq, "octDate": cur_date,
        "asOf": long(as_of), "asOfShort": short(as_of), "lastGF": long(last_gf),
        "built": dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d%H%M%S"),
    }
    if skipped:
        print("WARNING: these rows in sheet 'EAGR & GF' were NOT included in the dashboard:")
        for x in skipped:
            print("   -", x)
    else:
        print("All receipt rows were read.")
    print("Latest GF receipts read:")
    for g in GF[-3:]:
        print(f"   {g[0]}  {PG[g[1]]}  US${g[2]:,.0f}")
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
    ap.add_argument("--data-only", action="store_true",
                    help="only write the data file (no website files); used for a ready-made site")
    ap.add_argument("--out", default="data.json", help="output file when --data-only is used")
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
    if a.data_only:
        Path(a.out).write_text(json.dumps(out, separators=(",", ":")))
        print(f"Wrote {a.out}: {len(data['GF'])} GF receipts, {len(data['FC'])} FCDO receipts, "
              f"as of {data['F']['asOf']}")
        return
    site = Path(a.site)
    site.mkdir(exist_ok=True)
    for f in ("index.html", "app.js", "encrypt.html"):
        shutil.copy(Path("web") / f, site / f)
    (site / ".nojekyll").write_text("")
    (site / "data.json").write_text(json.dumps(out, separators=(",", ":")))
    print(f"Built {site}/ : {len(data['GF'])} GF receipts, {len(data['FC'])} FCDO receipts, "
          f"as of {data['F']['asOf']}")


if __name__ == "__main__":
    main()
