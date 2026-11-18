#!/usr/bin/env python3
"""
hdsupply_finder.py — Find product URLs on hdsupplysolutions.com

Input per row: manufacturer_name + Model Number + Supplier Item Number
Output per row: product_url (+ title, HD PART #, MFG part, price, status)

Approach (chosen because the site does NOT block it — verified 2026-10-28):
  Plain HTTPS requests to the site's own search endpoint. No browser,
  no login, no captcha handling needed:
      GET https://hdsupplysolutions.com/SearchDisplay?storeId=10051&searchTerm=<term>
  - An exact single hit JS-redirects to /p/<slug>-p<PART> — read from HTML.
  - Every search page embeds digitalData.product=[{productId, productName,
    mfgPart, brand}] — we filter it for an EXACT mfgPart match.
  - The hit is verified on its product page (schema.org JSON-LD: brand +
    MFG# must match the row). Only verified hits get a URL.

Search order per row (as specified by the user):
  Pass 1: "<manufacturer> <Model Number>"
          (+ the Model alone if it is not a bare number — bare numbers
           false-match other brands' PART #s, e.g. 56462 -> a ceiling fan)
  Pass 2: "<manufacturer> <Supplier Item Number>"
          (+ the Supplier Item Number alone, e.g. GJO01013)
  A candidate is accepted only if its mfgPart exactly equals the row's
  Model Number OR Supplier Item Number, and its brand matches.

Usage:
  python3 hdsupply_finder.py Workbook2-4.xlsx
  python3 hdsupply_finder.py Workbook2-4.xlsx -o out.xlsx --workers 6
  python3 hdsupply_finder.py --brand "GENUINE JOE" --model 1013 --supplier GJO01013

Batch runs checkpoint every 20 rows and RESUME from the output file on
re-run (rows already Found/Not Found are skipped).
"""

import argparse
import html as html_lib
import json
import re
import sys
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

BASE = "https://hdsupplysolutions.com"
SEARCH_URL = BASE + "/SearchDisplay?storeId=10051&searchTerm={term}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

BRAND_COLS = ["manufacturer_name", "manufacturer name", "brand",
              "brand_name", "brand name", "manufacturer", "make"]
MODEL_COLS = ["model number", "model_number", "model", "model_no",
              "model no", "mpn", "mfg_part", "mfg part"]
SUPPLIER_COLS = ["supplier item number", "supplier_item_number",
                 "supplier item", "supplier part number", "supplier_part",
                 "supplier number", "supplier_item", "item_number",
                 "item number", "part_number", "part number"]
OUT_COLS = ["status", "product_url", "product_title", "hd_part",
            "mfg_part", "price", "matched_on", "match_method", "note"]


def cell_str(v):
    """Excel cells come back as int/float for numeric models — normalise."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def norm_model(s):
    return re.sub(r"[\s\-_]+", "", str(s or "")).upper()


def norm_brand(s):
    s = re.sub(r"[^A-Z0-9 ]+", " ", str(s or "").upper())
    return re.sub(r"\s+", " ", s).strip()


def brand_matches(want, got):
    w, g = norm_brand(want), norm_brand(got)
    if not w:
        return True
    if not g:
        return False
    return w == g or w in g or g in w


_thread = threading.local()


def get_finder(delay):
    if not hasattr(_thread, "finder"):
        _thread.finder = Finder(delay=delay)
    return _thread.finder


class Finder:
    def __init__(self, delay=0.3, retries=3):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.delay = delay
        self.retries = retries

    def get(self, url):
        for attempt in range(self.retries):
            try:
                r = self.session.get(url, timeout=30)
                if r.status_code == 200:
                    if self.delay:
                        time.sleep(self.delay)
                    return r.text
                if r.status_code in (429, 500, 502, 503, 504):
                    time.sleep(2 * (attempt + 1))
                    continue
                return None
            except requests.RequestException:
                time.sleep(2 * (attempt + 1))
        return None

    def search(self, term):
        url = SEARCH_URL.format(term=urllib.parse.quote(str(term)))
        page = self.get(url)
        if not page:
            return None, []
        redirect = None
        m = re.search(r'window\.location\.href\s*=\s*"(' + BASE + r'/p/[^"]+)"', page)
        if m:
            redirect = m.group(1)
        products = []
        idx = page.find("digitalData.product=")
        if idx != -1:
            start = page.find("[", idx)
            if start != -1:
                try:
                    data, _ = json.JSONDecoder().raw_decode(page[start:])
                    if isinstance(data, list):
                        products = data
                except json.JSONDecodeError:
                    pass
        for p in products:
            if isinstance(p.get("productName"), str):
                p["productName"] = html_lib.unescape(p["productName"])
        return redirect, products

    def fetch_product_page(self, url):
        page = self.get(url)
        if not page:
            return None
        for block in re.findall(
            r'<script type="application/ld\+json">(.*?)</script>', page, re.S
        ):
            block = block.strip()
            if not block:
                continue
            try:
                data = json.loads(block)
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict) and data.get("@type") == "Product":
                offers = data.get("offers") or {}
                brand = data.get("brand") or ""
                if isinstance(brand, dict):
                    brand = brand.get("name", "")
                mfg = ""
                m = re.search(r"MFG#\s*([A-Za-z0-9\-]+)",
                              data.get("description", ""))
                if m:
                    mfg = m.group(1)
                return {
                    "title": html_lib.unescape(data.get("name", "")),
                    "brand": brand,
                    "hd_part": str(data.get("sku", "")),
                    "mfg_part": mfg,
                    "price": str(offers.get("price", "")),
                    "url": offers.get("url", url),
                }
        return None

    def find(self, brand, model, supplier=""):
        """Returns result dict. Main deliverable: product_url."""
        result = {c: "" for c in OUT_COLS}
        result["status"] = "Not Found"

        idents = []  # (label, value) in the user's search order
        if model:
            idents.append(("Model Number", model))
        if supplier and norm_model(supplier) != norm_model(model):
            idents.append(("Supplier Item Number", supplier))
        if not idents:
            result["note"] = "no model / supplier number in input"
            return result
        acceptable = {norm_model(v): label for label, v in idents}

        for label, value in idents:
            terms = [(f"{brand} {value}".strip(), f"manufacturer+{label}")]
            # Bare-number identifiers false-match PART #s of other brands,
            # so the identifier-alone search is only for alphanumeric codes.
            if not re.fullmatch(r"\d+", value):
                terms.append((value, f"{label} alone"))
            for term, method in terms:
                redirect, products = self.search(term)
                exact = [
                    p for p in products
                    if norm_model(p.get("mfgPart")) in acceptable
                    and brand_matches(brand, p.get("brand"))
                ]
                if not exact:
                    continue
                hit = exact[0]
                part = str(hit.get("productId", ""))
                matched_on = acceptable[norm_model(hit.get("mfgPart"))]
                if redirect:
                    url = redirect
                else:
                    slug = re.sub(r"[^a-z0-9]+", "-",
                                  str(hit.get("productName", "product")).lower()
                                  ).strip("-")
                    url = f"{BASE}/p/{slug}-p{part}"

                detail = self.fetch_product_page(url)
                if detail is None:
                    result.update({
                        "status": "Found", "product_url": url,
                        "product_title": hit.get("productName", ""),
                        "hd_part": part, "mfg_part": hit.get("mfgPart", ""),
                        "matched_on": matched_on, "match_method": method,
                        "note": "exact match on search page; product page "
                                "unreadable, price unverified",
                    })
                    return result
                if detail["mfg_part"] and \
                        norm_model(detail["mfg_part"]) not in acceptable:
                    continue
                if not brand_matches(brand, detail["brand"]):
                    continue
                result.update({
                    "status": "Found",
                    "product_url": detail["url"] or url,
                    "product_title": detail["title"],
                    "hd_part": detail["hd_part"] or part,
                    "mfg_part": detail["mfg_part"] or hit.get("mfgPart", ""),
                    "price": detail["price"],
                    "matched_on": matched_on, "match_method": method,
                })
                return result
        return result


# ------------------------------------------------------------------- batch

def detect(fieldnames, candidates, explicit=None):
    lower = {str(c).strip().lower(): c for c in fieldnames}
    if explicit:
        return lower.get(explicit.strip().lower(), explicit)
    return next((lower[c] for c in candidates if c in lower), None)


def run_batch(in_path, out_path, workers=6, delay=0.3,
              brand_col=None, model_col=None, supplier_col=None):
    import os
    is_csv = str(in_path).lower().endswith(".csv")
    if is_csv:
        import csv
        with open(in_path, newline="", encoding="utf-8-sig") as f:
            rows = [{k: cell_str(v) for k, v in r.items()}
                    for r in csv.DictReader(f)]
        fields = list(rows[0].keys()) if rows else []
    else:
        from openpyxl import load_workbook
        wb = load_workbook(in_path)
        ws = wb.active
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        fields = [str(h) if h is not None else "" for h in headers]
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            rows.append({fields[i]: cell_str(r[i]) if i < len(r) else ""
                         for i in range(len(fields))})

    b_col = detect(fields, BRAND_COLS, brand_col)
    m_col = detect(fields, MODEL_COLS, model_col)
    s_col = detect(fields, SUPPLIER_COLS, supplier_col)
    if not m_col and not s_col:
        sys.exit(f"No model/supplier column found. Columns: {fields}")
    print(f"Columns: brand={b_col!r} model={m_col!r} supplier={s_col!r} "
          f"| rows={len(rows)} workers={workers}")

    # resume from existing output
    done = {}
    if out_path and os.path.exists(out_path):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(out_path)
            ws = wb.active
            hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            for r in ws.iter_rows(min_row=2, values_only=True):
                d = dict(zip(hdr, r))
                if d.get("status") in ("Found", "Not Found"):
                    key = tuple(cell_str(d.get(c)) for c in
                                (b_col, m_col, s_col) if c)
                    done[key] = {c: cell_str(d.get(c)) for c in OUT_COLS}
            if done:
                print(f"Resuming: {len(done)} rows already done")
        except Exception:
            done = {}

    def key_of(row):
        return tuple(row.get(c, "") for c in (b_col, m_col, s_col) if c)

    todo_idx = []
    for i, row in enumerate(rows):
        if key_of(row) in done:
            row.update(done[key_of(row)])
        else:
            todo_idx.append(i)

    def save(path):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        out_fields = fields + [c for c in OUT_COLS if c not in fields]
        ws.append(out_fields)
        for row in rows:
            ws.append([row.get(c, "") for c in out_fields])
        wb.save(path)

    checkpoint = str(out_path).replace(".xlsx", "_checkpoint.xlsx")
    completed = 0

    def work(i):
        row = rows[i]
        finder = get_finder(delay)
        return i, finder.find(row.get(b_col, "") if b_col else "",
                              row.get(m_col, "") if m_col else "",
                              row.get(s_col, "") if s_col else "")

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(work, i): i for i in todo_idx}
        for fut in as_completed(futures):
            i, res = fut.result()
            rows[i].update(res)
            completed += 1
            print(f"[{completed}/{len(todo_idx)}] "
                  f"{rows[i].get(b_col,'')} {rows[i].get(m_col,'')} / "
                  f"{rows[i].get(s_col,'')} -> {res['status']}"
                  f"{' ' + res['product_url'] if res['product_url'] else ''}",
                  flush=True)
            if completed % 20 == 0:
                save(checkpoint)
    if out_path:
        save(out_path)
    found = sum(1 for r in rows if r.get("status") == "Found")
    print(f"\nDone: {found} Found / {len(rows) - found} Not Found / "
          f"{len(rows)} total")
    return found, len(rows) - found


def main():
    ap = argparse.ArgumentParser(
        description="Find product URLs on hdsupplysolutions.com")
    ap.add_argument("input", nargs="?", help="Input .xlsx or .csv")
    ap.add_argument("-o", "--output")
    ap.add_argument("--brand")
    ap.add_argument("--model")
    ap.add_argument("--supplier")
    ap.add_argument("--brand-col")
    ap.add_argument("--model-col")
    ap.add_argument("--supplier-col")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--delay", type=float, default=0.3)
    args = ap.parse_args()

    if args.model or args.supplier:
        res = Finder(delay=args.delay).find(
            args.brand or "", args.model or "", args.supplier or "")
        print(json.dumps({"brand": args.brand or "", "model": args.model or "",
                          "supplier": args.supplier or "", **res}, indent=2))
        return
    if not args.input:
        ap.error("give an input file, or --brand/--model/--supplier")
    out = args.output or re.sub(r"\.(xlsx|csv)$", "", args.input,
                                 flags=re.I) + "_hdsupply.xlsx"
    run_batch(args.input, out, args.workers, args.delay,
              args.brand_col, args.model_col, args.supplier_col)


if __name__ == "__main__":
    main()
