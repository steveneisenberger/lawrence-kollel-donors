#!/usr/bin/env python3
"""
Build the encrypted donor dataset for the Lawrence Kollel donor site.

Usage:
    python3 build.py "2026 Donations.xlsx" --password "YourPassword"

Reads the spreadsheet, extracts donors / pledges / donations, encrypts the
result with the password (PBKDF2-SHA256 + AES-256-GCM) and writes
site/data.enc.json. The plaintext never leaves this machine; only the
encrypted file is committed.
"""
import argparse, base64, json, os, re, sys
from datetime import date

import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

PBKDF2_ITERS = 250_000
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def key(last, first):
    return (norm(last).lower(), norm(first).lower())


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("$", ""))
    except ValueError:
        return 0.0


def read_contacts(wb):
    ws = wb["Sheet1"]
    contacts = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        last, first, title, phone, email, addr, city, state, zip_ = (list(row) + [None] * 9)[:9]
        if not norm(last):
            continue
        k = key(last, first)
        c = contacts.setdefault(k, {})
        for field, val in (("title", title), ("phone", phone), ("email", email),
                           ("address", addr), ("city", city), ("state", state), ("zip", zip_)):
            val = norm(val)
            if field == "zip" and val.endswith(".0"):
                val = val[:-2]
            if field == "phone" and val.endswith(".0"):
                val = val[:-2]
            if val and not c.get(field):
                c[field] = val
    return contacts


def read_contributions(wb, year):
    ws = wb["Contribution 2026"]
    # Locate monthly columns from the header row (row 2).
    month_cols = {}
    header = {c: norm(ws.cell(2, c).value) for c in range(1, ws.max_column + 1)}
    for c, h in header.items():
        if h in MONTHS:
            month_cols[h] = c
    col = {h: c for c, h in header.items() if h}
    c_last, c_first = 2, 3
    c_out, c_pledge = col["Outstanding 2025"], col["2026"]
    c_balance, c_don, c_total = col["Balance"], col["Donations"], col["Total Donations"]
    c_pm = col["2026 Parlor Meeting"]  # header sits over the amount column; name cols are the two before it
    c_pm_last, c_pm_first, c_pm_amt, c_pm_method = c_pm - 2, c_pm - 1, c_pm, c_pm + 1

    donors = {}
    order = []

    def get(last, first):
        k = key(last, first)
        if k not in donors:
            donors[k] = {"last": norm(last), "first": norm(first), "donations": [],
                         "pledge_outstanding_prior": 0.0, "pledge_current": 0.0,
                         "pledge_balance": 0.0}
            order.append(k)
        return donors[k]

    for r in range(3, ws.max_row + 1):
        last = ws.cell(r, c_last).value
        if norm(last):
            d = get(last, ws.cell(r, c_first).value)
            d["pledge_outstanding_prior"] = num(ws.cell(r, c_out).value)
            d["pledge_current"] = num(ws.cell(r, c_pledge).value)
            d["pledge_balance"] = num(ws.cell(r, c_balance).value)
            for mi, mname in enumerate(MONTHS, start=1):
                c = month_cols.get(mname)
                if not c:
                    continue
                amt = num(ws.cell(r, c).value)
                method = norm(ws.cell(r, c + 1).value) if c + 1 <= ws.max_column else ""
                # Nov/Dec have no method column in this layout; guard against reading a month name.
                if method in MONTHS:
                    method = ""
                if amt:
                    d["donations"].append({
                        "type": "donation",
                        "date": f"{year}-{mi:02d}",   # month precision only
                        "amount": round(abs(amt), 2),
                        "method": method,
                    })
        # Parlor meeting block (separate name columns, may not align with the row's donor).
        pm_last = ws.cell(r, c_pm_last).value
        if norm(pm_last):
            amt = num(ws.cell(r, c_pm_amt).value)
            if amt:
                d = get(pm_last, ws.cell(r, c_pm_first).value)
                d["donations"].append({
                    "type": "parlor",
                    "date": None,
                    "event": f"{year} Parlor Meeting",
                    "amount": round(abs(amt), 2),
                    "method": norm(ws.cell(r, c_pm_method).value),
                })
    return [donors[k] for k in order]


def build(xlsx, year):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    contacts = read_contacts(wb)
    donors = read_contributions(wb, year)
    out = []
    for i, d in enumerate(donors):
        c = contacts.get(key(d["last"], d["first"]), {})
        don_total = sum(x["amount"] for x in d["donations"] if x["type"] == "donation")
        pm_total = sum(x["amount"] for x in d["donations"] if x["type"] == "parlor")
        out.append({
            "id": i + 1,
            "last": d["last"], "first": d["first"],
            "title": c.get("title", ""), "phone": c.get("phone", ""), "email": c.get("email", ""),
            "address": c.get("address", ""), "city": c.get("city", ""),
            "state": c.get("state", ""), "zip": c.get("zip", ""),
            "pledge_outstanding_prior": d["pledge_outstanding_prior"],
            "pledge_current": d["pledge_current"],
            "pledge_balance": d["pledge_balance"],
            "donations_total": round(don_total, 2),
            "parlor_total": round(pm_total, 2),
            "total": round(don_total + pm_total, 2),
            "donations": d["donations"],
        })
    return {"year": year, "generated": date.today().isoformat(),
            "source": os.path.basename(xlsx), "donors": out}


def encrypt(payload: dict, password: str) -> dict:
    salt = os.urandom(16)
    iv = os.urandom(12)
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=PBKDF2_ITERS)
    k = kdf.derive(password.encode("utf-8"))
    ct = AESGCM(k).encrypt(iv, json.dumps(payload, separators=(",", ":")).encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"kdf": "PBKDF2-SHA256", "iterations": PBKDF2_ITERS, "cipher": "AES-256-GCM",
            "salt": b64(salt), "iv": b64(iv), "data": b64(ct)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--password", required=True)
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__) or ".", "docs", "data.enc.json"))
    ap.add_argument("--dump-plain", help="(debug) also write plaintext JSON here — never commit it")
    a = ap.parse_args()
    payload = build(a.xlsx, a.year)
    if a.dump_plain:
        with open(a.dump_plain, "w") as f:
            json.dump(payload, f, indent=1)
    with open(a.out, "w") as f:
        json.dump(encrypt(payload, a.password), f)
    n = len(payload["donors"])
    nd = sum(len(d["donations"]) for d in payload["donors"])
    tot = sum(d["total"] for d in payload["donors"])
    print(f"{n} donors, {nd} gifts, ${tot:,.2f} total -> {a.out}")


if __name__ == "__main__":
    main()
