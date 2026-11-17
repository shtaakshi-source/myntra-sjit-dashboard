#!/usr/bin/env python3
"""
Nat Habit × Myntra — SJIT (FBM) Intelligence Dashboard builder.

Drop the latest Myntra "Seller Orders Report" CSV into ./data and run:

    python build_dashboard.py

It reads the newest report, keeps only orders whose cancellation / RTO / lost /
return fields are all blank, and writes a self-contained dashboard to
./output/myntra_sjit_dashboard.html (open it in any browser).

Options:
    --report PATH     use a specific report instead of the newest file in ./data
    --combine         merge every report in ./data (de-duplicated on order line id,
                      newest file wins) — useful for building a full month
    --out PATH        where to write the HTML (default: output/myntra_sjit_dashboard.html)
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
CONFIG_DIR = ROOT / "config"
TEMPLATE = ROOT / "template" / "dashboard_template.html"
DEFAULT_OUT = ROOT / "output" / "myntra_sjit_dashboard.html"

WAREHOUSE_MAP = CONFIG_DIR / "warehouse_id_mapping.xlsx"   # .csv also accepted
SKU_MAP = CONFIG_DIR / "sku_mapping.json"
CATEGORY_MAP = CONFIG_DIR / "category_mapping.xlsx"     # Style ID → nh_sku, category (.csv also accepted)
FRESH_CFG = CONFIG_DIR / "fresh_skus.json"

# Columns that must ALL be blank for an order to count
EXCLUSION_COLS = {
    "C": ["cancellation reason id fk", "cancellation reason", "cancelled on"],  # cancelled
    "R": ["rto creation date"],                                               # RTO
    "L": ["lost date"],                                                       # lost
    "T": ["return creation date"],                                            # returned
}
REQUIRED_COLS = ["po_type", "seller sku code", "created on", "warehouse id"]


# ─────────────────────────── loading ───────────────────────────
def find_reports() -> list[Path]:
    files = [p for p in DATA_DIR.glob("*") if p.suffix.lower() in (".csv", ".xlsx", ".xls")]
    return sorted(files, key=lambda p: p.stat().st_mtime)


def read_report(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [c.strip().lower() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        sys.exit(f"✗ {path.name} is missing column(s): {', '.join(missing)}. "
                 "Is this the Myntra Seller Orders Report?")
    df["__file"] = path.name
    return df


def read_warehouses() -> dict[str, list[str]]:
    path = WAREHOUSE_MAP
    if not path.exists():
        alt = path.with_suffix(".csv")
        if not alt.exists():
            print("! No warehouse mapping found in config/ — warehouses will show as IDs.")
            return {}
        path = alt
    wh = pd.read_csv(path, dtype=str) if path.suffix == ".csv" else pd.read_excel(path, dtype=str)
    wh.columns = [c.strip().lower() for c in wh.columns]
    out = {}
    for _, r in wh.iterrows():
        try:
            wid = str(int(float(r["warehouse id"])))
        except (ValueError, TypeError):
            continue
        out[wid] = [str(r.get("warehouse name", "")).strip(), str(r.get("state", "")).strip()]
    return out


def read_categories() -> dict[str, str]:
    """Style ID → NH category. Styles not in the file fall back to Myntra's article type."""
    path = CATEGORY_MAP
    if not path.exists():
        alt = path.with_suffix(".csv")
        if not alt.exists():
            print("! No category mapping found in config/ — using Myntra article type for categories.")
            return {}
        path = alt
    cm = pd.read_csv(path, dtype=str) if path.suffix == ".csv" else pd.read_excel(path, dtype=str)
    cm.columns = [c.strip().lower() for c in cm.columns]
    if "style id" not in cm.columns or "category" not in cm.columns:
        sys.exit(f"✗ {path.name} needs 'Style ID' and 'category' columns.")
    cm = cm.dropna(subset=["style id", "category"])
    return {to_int_id(s): str(c).strip() for s, c in zip(cm["style id"], cm["category"]) if str(c).strip()}


# ─────────────────────────── processing ───────────────────────────
def exclusion_reason(df: pd.DataFrame) -> pd.Series:
    reason = pd.Series("", index=df.index)
    # Assign in reverse priority so cancellation wins when several are filled
    for code in ["T", "L", "R", "C"]:
        cols = [c for c in EXCLUSION_COLS[code] if c in df.columns]
        if not cols:
            continue
        filled = df[cols].apply(lambda s: s.str.strip().ne("")).any(axis=1)
        reason[filled] = code
    return reason


def parse_date(s: str) -> str:
    s = (s or "").strip()[:10]
    if len(s) == 10 and s[2] in "-/" and s[5] in "-/":   # DD-MM-YYYY
        return f"{s[6:10]}-{s[3:5]}-{s[0:2]}"
    return s


def to_int_id(v: str) -> str:
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        return str(v).strip()


def build_payload(df: pd.DataFrame, file_label: str, catmap: dict[str, str]) -> dict:
    df = df.copy()
    df["__x"] = exclusion_reason(df)
    df["__d"] = df["created on"].map(parse_date)
    df["__amt"] = pd.to_numeric(df.get("final amount", 0), errors="coerce").fillna(0).round().astype(int)
    for col in ["style name", "article type", "state", "style id"]:
        if col not in df.columns:
            df[col] = ""

    dates = sorted(d for d in df["__d"].unique() if d)
    d_idx = {d: i for i, d in enumerate(dates)}
    sku_idx, skus, st_idx, states, rows = {}, [], {}, [], []

    sku_col = df["seller sku code"].str.strip().tolist()
    name_col = df["style name"].tolist()
    at_col = df["article type"].tolist()
    style_col = df["style id"].map(to_int_id).tolist()
    po_col = df["po_type"].str.strip().str.upper().tolist()
    wh_col = df["warehouse id"].map(to_int_id).tolist()
    st_col = df["state"].str.strip().str.upper().tolist()

    d_col = df["__d"].tolist()
    amt_col = df["__amt"].tolist()
    x_col = df["__x"].tolist()

    for i in range(len(df)):
        d = d_col[i]
        if not d:
            continue
        k = sku_col[i]
        if k not in sku_idx:
            sku_idx[k] = len(skus)
            skus.append([k, name_col[i], catmap.get(style_col[i]) or at_col[i] or "Unknown", at_col[i]])
        s = st_col[i]
        if s not in st_idx:
            st_idx[s] = len(states)
            states.append(s)
        rows.append([d_idx[d], sku_idx[k], 1 if po_col[i] == "FBM" else 0,
                     wh_col[i], st_idx[s], int(amt_col[i]), x_col[i]])

    return {"file": file_label, "dates": dates, "skus": skus, "states": states, "rows": rows}


def print_summary(df: pd.DataFrame) -> None:
    x = exclusion_reason(df)
    valid = df[x.eq("")]
    fbm = valid["po_type"].str.upper().eq("FBM").sum()
    print(f"  Order lines in file : {len(df):,}")
    print(f"  Excluded            : {(x == 'C').sum():,} cancelled · {(x == 'R').sum():,} RTO · "
          f"{(x == 'T').sum():,} returned · {(x == 'L').sum():,} lost")
    print(f"  Counted             : {len(valid):,}")
    if len(valid):
        print(f"  SJIT (FBM) share    : {fbm:,} orders = {fbm / len(valid) * 100:.1f}% of all counted orders")


# ─────────────────────────── main ───────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser(description="Build the Myntra SJIT dashboard from a Seller Orders Report.")
    ap.add_argument("--report", type=Path, help="specific report file to use")
    ap.add_argument("--combine", action="store_true", help="merge all reports in ./data")
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT, help="output HTML path")
    args = ap.parse_args()

    if args.report:
        if not args.report.exists():
            sys.exit(f"✗ Report not found: {args.report}")
        files = [args.report]
    else:
        files = find_reports()
        if not files:
            sys.exit(f"✗ No report found. Put the Myntra Seller Orders Report CSV in {DATA_DIR}/")
        if not args.combine:
            files = files[-1:]

    print("Reading:")
    frames = []
    for f in files:
        print(f"  • {f.name}")
        frames.append(read_report(f))
    df = pd.concat(frames, ignore_index=True)
    if len(frames) > 1 and "order line id" in df.columns:
        before = len(df)
        df = df.drop_duplicates(subset="order line id", keep="last")  # newest file wins
        print(f"  Merged {len(frames)} files, removed {before - len(df):,} duplicate order lines")

    label = files[0].name if len(files) == 1 else f"{len(files)} reports combined"
    catmap = read_categories()
    payload = build_payload(df, label, catmap)
    if not payload["rows"]:
        sys.exit("✗ No rows with a valid 'created on' date.")

    print("\nSummary:")
    print_summary(df)
    valid = df[exclusion_reason(df).eq("")]
    unmapped = valid[~valid["style id"].map(to_int_id).isin(catmap.keys())]
    print(f"  Category mapping    : {len(valid) - len(unmapped):,} of {len(valid):,} counted orders mapped")
    if len(unmapped):
        print("  Unmapped styles (add to config/category_mapping.xlsx; article type used for now):")
        top = (unmapped.groupby(["style id", "seller sku code"]).size()
               .sort_values(ascending=False).head(10))
        for (sid, sku), n in top.items():
            print(f"    {sid:>10}  {sku:<22} {n:>5} orders")

    d0, d1 = payload["dates"][0], payload["dates"][-1]
    fmt = lambda d: datetime.strptime(d, "%Y-%m-%d").strftime("%-d %b %Y") if sys.platform != "win32" \
        else datetime.strptime(d, "%Y-%m-%d").strftime("%#d %b %Y")
    rng = f"{fmt(d0)} – {fmt(d1)}"

    def js(obj) -> str:
        # compact JSON, safe inside a <script> tag
        return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    html = TEMPLATE.read_text(encoding="utf-8")
    html = (html
            .replace("__SMAP__", js(json.loads(SKU_MAP.read_text(encoding="utf-8"))))
            .replace("__CATMAP__", js(catmap))
            .replace("__WH__", js(read_warehouses()))
            .replace("__FRESH__", js(json.loads(FRESH_CFG.read_text(encoding="utf-8"))))
            .replace("__BAKED_RANGE__", rng)
            .replace("__BUILT_AT__", datetime.now().strftime("%d %b %Y, %H:%M"))
            .replace("__BAKED__", js(payload)))  # last: the data itself may contain placeholder-like text

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(html, encoding="utf-8")
    print(f"\n✓ Dashboard written to {args.out}  ({args.out.stat().st_size / 1024:.0f} KB, {rng})")


if __name__ == "__main__":
    main()
