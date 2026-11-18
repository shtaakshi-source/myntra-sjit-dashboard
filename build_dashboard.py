#!/usr/bin/env python3
"""
Nat Habit × Myntra — SJIT (FBM) Intelligence Dashboard builder.

Daily routine
-------------
1. Drop the day's Myntra **Seller Orders Report** CSV into ./data
   (and the **Seller Inventory Report** CSV whenever you pull one).
2. Run:   python build_dashboard.py
3. Open:  output/myntra_sjit_dashboard.html

What the script does
--------------------
* Merges EVERY Seller Orders Report in ./data, de-duplicated on order line id
  (the newest file wins, so later cancellations/returns overwrite older rows).
  Keep old reports in ./data: that is how the history / trend view grows.
* Uses the NEWEST Seller Inventory Report in ./data for FC stock.
* Counts an order only if cancellation / RTO / lost / return fields are all blank.
* Maps every style ID to NH SKU, category and shelf life using
  config/MPM_Myntra.xlsx (the Myntra sheet: style ID = col B, NH SKU = col C,
  shelf life = col J, category = col U).
* Lists styles that sold but are missing from the MPM (printed here and
  saved to output/mpm_gaps.csv) so you can update the MPM sheet.

Options
-------
    --latest-only    use only the newest orders report instead of merging all
    --out PATH       where to write the HTML (default output/myntra_sjit_dashboard.html)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
CONFIG_DIR = ROOT / "config"
OUT_DIR = ROOT / "output"
TEMPLATE = ROOT / "template" / "dashboard_template.html"
DEFAULT_OUT = OUT_DIR / "myntra_sjit_dashboard.html"

MPM_FILE = CONFIG_DIR / "MPM_Myntra.xlsx"
WAREHOUSE_MAP = CONFIG_DIR / "warehouse_id_mapping.xlsx"
NETWORK_CFG = CONFIG_DIR / "myntra_network.json"

EXCLUSION_COLS = {
    "C": ["cancellation reason id fk", "cancellation reason", "cancelled on"],
    "R": ["rto creation date"],
    "L": ["lost date"],
    "T": ["return creation date"],
}
ORDER_COLS = ["po_type", "seller sku code", "created on", "warehouse id", "style id"]
INV_COLS = ["warehouse id", "po_type", "style id", "inventory count"]

# MPM columns by name, with the column letter as a fallback
MPM_COLS = {"style": ("channel_sr_no_1", "B"), "nh_sku": ("nh_sku", "C"),
            "shelf": ("shelf_life", "J"), "cat": ("infytrix_sub_category", "U"),
            "status": ("status", "K")}


# ─────────────────────────── helpers ───────────────────────────
def to_id(v) -> str:
    s = str(v).strip()
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def file_date(p: Path) -> str:
    """Date in the file name (YYYY-MM-DD) or the file's modified time, for ordering."""
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", p.name)
    return (m.group(1) if m else datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")) + \
        f"_{p.stat().st_mtime:.0f}"


def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def classify_reports():
    orders, inventory = [], []
    for p in sorted(DATA_DIR.glob("*")):
        if p.suffix.lower() not in (".csv", ".xlsx", ".xls") or p.name.startswith("~$"):
            continue
        try:
            head = read_any(p).columns if p.suffix.lower() != ".csv" else \
                [c.strip().lower() for c in pd.read_csv(p, nrows=0).columns]
        except Exception as e:  # noqa: BLE001
            print(f"  ! Skipping {p.name}: {e}")
            continue
        if all(c in head for c in ORDER_COLS):
            orders.append(p)
        elif all(c in head for c in INV_COLS):
            inventory.append(p)
        else:
            print(f"  ! Skipping {p.name}: not a Seller Orders or Seller Inventory report")
    return sorted(orders, key=file_date), sorted(inventory, key=file_date)


def exclusion_reason(df: pd.DataFrame) -> pd.Series:
    reason = pd.Series("", index=df.index)
    for code in ["T", "L", "R", "C"]:  # cancellation wins when several are filled
        cols = [c for c in EXCLUSION_COLS[code] if c in df.columns]
        if cols:
            reason[df[cols].apply(lambda s: s.str.strip().ne("")).any(axis=1)] = code
    return reason


# Same place, different spellings in Myntra exports
STATE_ALIAS = {"TS": "TG", "CG": "CT", "UK": "UT", "OD": "OR"}
CITY_ALIAS = {"Bengaluru": "Bangalore", "Bangalore Urban": "Bangalore", "Gurugram": "Gurgaon",
              "Bombay": "Mumbai", "New Delhi": "Delhi", "Calcutta": "Kolkata", "Madras": "Chennai"}


def parse_date(s: str) -> str:
    s = (s or "").strip()[:10]
    if len(s) == 10 and s[2] in "-/" and s[5] in "-/":
        return f"{s[6:10]}-{s[3:5]}-{s[0:2]}"
    return s


# ─────────────────────────── config ───────────────────────────
def read_mpm() -> dict[str, dict]:
    if not MPM_FILE.exists():
        sys.exit(f"✗ MPM not found: {MPM_FILE}. Save the MPM workbook there as MPM_Myntra.xlsx.")
    sheets = pd.read_excel(MPM_FILE, sheet_name=None, dtype=str)
    # Prefer the sheet whose 'channel' column is Myntra (Sheet2 in the current workbook)
    best, best_score = None, -1
    for name, df in sheets.items():
        cols = [str(c).strip().lower() for c in df.columns]
        score = 0
        if "channel" in cols:
            ch = df.iloc[:, cols.index("channel")].astype(str).str.strip().str.lower()
            score = (ch == "myntra").mean()
        elif name.strip().lower() == "sheet2":
            score = 0.5
        if score > best_score:
            best, best_score = name, score
    df = sheets[best]
    df.columns = [str(c).strip().lower() for c in df.columns]
    letters = {chr(65 + i): c for i, c in enumerate(df.columns)}

    def col(key):
        name, letter = MPM_COLS[key]
        if name in df.columns:
            return df[name]
        return df[letters[letter]] if letter in letters else pd.Series("", index=df.index)

    mpm = {}
    for style, nh, shelf, cat, status in zip(col("style"), col("nh_sku"), col("shelf"), col("cat"), col("status")):
        sid = to_id(style)
        if not sid or sid == "nan":
            continue
        try:
            sl = int(float(shelf))
        except (ValueError, TypeError):
            sl = None
        mpm[sid] = {"nh": str(nh).strip() if str(nh) != "nan" else "",
                    "shelf": sl,
                    "cat": str(cat).strip() if str(cat) != "nan" else "",
                    "status": str(status).strip() if str(status) != "nan" else ""}
    print(f"  MPM: sheet '{best}', {len(mpm):,} Myntra style IDs")
    return mpm


def read_warehouses() -> dict[str, list[str]]:
    if not WAREHOUSE_MAP.exists():
        print("  ! No warehouse mapping in config/ — warehouses will show as IDs.")
        return {}
    wh = pd.read_excel(WAREHOUSE_MAP, dtype=str)
    wh.columns = [c.strip().lower() for c in wh.columns]
    wh = wh.fillna("")
    return {to_id(r["warehouse id"]): [str(r.get("warehouse name", "")).strip(), str(r.get("state", "")).strip()]
            for _, r in wh.iterrows() if str(r.get("warehouse id", "")).strip()}


# ─────────────────────────── main build ───────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(description="Build the Myntra SJIT dashboard.")
    ap.add_argument("--latest-only", action="store_true", help="use only the newest orders report")
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = ap.parse_args()

    print("Config:")
    mpm = read_mpm()
    whmap = read_warehouses()
    network = json.loads(NETWORK_CFG.read_text(encoding="utf-8"))

    print("\nReports in data/:")
    order_files, inv_files = classify_reports()
    if not order_files:
        sys.exit(f"✗ No Seller Orders Report found in {DATA_DIR}/")
    if args.latest_only:
        order_files = order_files[-1:]
    for f in order_files:
        print(f"  • orders     {f.name}")
    for f in inv_files[-1:]:
        print(f"  • inventory  {f.name}  (newest only)")
    if not inv_files:
        print("  ! No Seller Inventory Report — FC stock will show as zero and send quantities ignore stock.")

    # ── orders
    df = pd.concat([read_any(f) for f in order_files], ignore_index=True)
    if len(order_files) > 1 and "order line id" in df.columns:
        n0 = len(df)
        df = df.drop_duplicates(subset="order line id", keep="last")
        print(f"  Merged {len(order_files)} order reports → {len(df):,} unique order lines ({n0 - len(df):,} duplicates dropped)")
    for c in ["style name", "article type", "state", "city", "final amount"]:
        if c not in df.columns:
            df[c] = ""
    df["__x"] = exclusion_reason(df)
    df["__d"] = df["created on"].map(parse_date)
    df = df[df["__d"].str.len() == 10].copy()
    df["__style"] = df["style id"].map(to_id)
    df["__sku"] = df["seller sku code"].str.strip()
    df["__amt"] = pd.to_numeric(df["final amount"], errors="coerce").fillna(0).round().astype(int)

    # ── inventory (newest report, SJIT rows only — PPMP rows are our own virtual stock)
    inv = pd.DataFrame(columns=["warehouse id", "style id", "seller sku code", "qty"])
    inv_file, inv_asof = "", ""
    if inv_files:
        inv_file = inv_files[-1].name
        raw = read_any(inv_files[-1])
        raw["qty"] = pd.to_numeric(raw["inventory count"], errors="coerce").fillna(0) - \
            pd.to_numeric(raw.get("blocked count", 0), errors="coerce").fillna(0)
        if "modified on" in raw.columns:
            inv_asof = max((parse_date(x) for x in raw["modified on"] if x), default="")
        inv = raw[raw["po_type"].str.upper().isin(["SJIT", "FBM"]) & (raw["qty"] > 0)].copy()
        inv["warehouse id"] = inv["warehouse id"].map(to_id)
        inv["style id"] = inv["style id"].map(to_id)
        inv["seller sku code"] = inv.get("seller sku code", "").astype(str).str.strip()
        unknown_inv_wh = sorted(set(raw["warehouse id"].map(to_id)) - set(whmap))
    else:
        unknown_inv_wh = []

    # ── product table, keyed by style ID
    products, pidx = [], {}

    def product(style, sku, name, at):
        if style in pidx:
            p = products[pidx[style]]
            if not p[3] and name:
                p[3] = name
            return pidx[style]
        m = mpm.get(style)
        pidx[style] = len(products)
        products.append([style, sku, m["nh"] if m else "", name or "", (m["cat"] if m and m["cat"] else at) or "Unknown",
                         m["shelf"] if m else None, m["status"] if m else "", 1 if m else 0, at or ""])
        return pidx[style]

    dates = sorted(df["__d"].unique())
    d_i = {d: i for i, d in enumerate(dates)}
    states, s_i, cities, c_i = [], {}, [], {}
    rows = []
    cols = zip(df["__d"], df["__style"], df["__sku"], df["style name"], df["article type"],
               df["po_type"].str.strip().str.upper(), df["warehouse id"].map(to_id),
               df["state"].str.strip().str.upper().replace(STATE_ALIAS), df["city"].str.strip().str.title().replace(CITY_ALIAS), df["__amt"], df["__x"])
    for d, style, sku, name, at, po, wh, st, city, amt, x in cols:
        p = product(style, sku, name, at)
        if st not in s_i:
            s_i[st] = len(states); states.append(st)
        if city not in c_i:
            c_i[city] = len(cities); cities.append(city)
        rows.append([d_i[d], p, 1 if po == "FBM" else 0, wh, s_i[st], c_i[city], int(amt), x])

    inv_rows = []
    for wh, style, sku, q in zip(inv["warehouse id"], inv["style id"], inv["seller sku code"], inv["qty"]):
        inv_rows.append([wh, product(style, sku, "", ""), int(q)])

    # ── MPM gaps report
    valid = df[df["__x"] == ""]
    gaps = (valid[~valid["__style"].isin(mpm)]
            .groupby(["__style", "__sku"])
            .agg(orders=("__d", "size"), last_order=("__d", "max"), name=("style name", "first"))
            .reset_index().sort_values("orders", ascending=False))
    inv_gaps = sorted(set(inv["style id"]) - set(mpm)) if len(inv) else []
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    gaps.rename(columns={"__style": "style id", "__sku": "seller sku code"}).to_csv(OUT_DIR / "mpm_gaps.csv", index=False)

    # ── summary
    x = df["__x"]
    fbm = valid["po_type"].str.upper().eq("FBM").sum()
    print("\nSummary:")
    print(f"  Period              : {dates[0]} → {dates[-1]} ({len(dates)} days with orders)")
    print(f"  Order lines         : {len(df):,}  (excluded {(x == 'C').sum():,} cancelled · {(x == 'R').sum():,} RTO · "
          f"{(x == 'T').sum():,} returned · {(x == 'L').sum():,} lost)")
    print(f"  Counted             : {len(valid):,}   SJIT share {fbm / max(1, len(valid)) * 100:.1f}%")
    print(f"  FC/FDC/M-Now stock  : {int(inv['qty'].sum()) if len(inv) else 0:,} units (as of {inv_asof or 'n/a'})")
    if len(gaps):
        print(f"\n  ⚠ {len(gaps)} style(s) sold but are NOT in the MPM ({int(gaps['orders'].sum())} orders). "
              f"Update config/MPM_Myntra.xlsx — full list in output/mpm_gaps.csv:")
        for _, r in gaps.head(15).iterrows():
            print(f"     {r['__style']:>10}  {r['__sku']:<22} {r['orders']:>5} orders  (last {r['last_order']})")
    else:
        print("  MPM coverage        : every style that sold is in the MPM ✓")
    if inv_gaps:
        print(f"  ⚠ {len(inv_gaps)} style(s) in Myntra stock are not in the MPM")
    unknown_wh = sorted((set(df["warehouse id"].map(to_id)) | set(unknown_inv_wh)) - set(whmap))
    if unknown_wh:
        print(f"  ⚠ Warehouse IDs not in config/warehouse_id_mapping.xlsx: {', '.join(unknown_wh)}")

    payload = {
        "files": [f.name for f in order_files], "invFile": inv_file, "invAsOf": inv_asof,
        "dates": dates, "products": products, "states": states, "cities": cities,
        "rows": rows, "inv": inv_rows, "invGapStyles": inv_gaps, "unknownWh": unknown_wh,
    }

    def js(obj) -> str:
        return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    html = TEMPLATE.read_text(encoding="utf-8")
    html = (html.replace("__WH__", js(whmap))
                .replace("__NETWORK__", js(network))
                .replace("__BUILT_AT__", datetime.now().strftime("%d %b %Y, %H:%M"))
                .replace("__DATA__", js(payload)))
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(html, encoding="utf-8")
    print(f"\n✓ Dashboard written to {args.out}  ({args.out.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
