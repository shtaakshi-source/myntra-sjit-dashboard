# Nat Habit × Myntra — SJIT (FBM) Intelligence Dashboard

Builds a self-contained HTML dashboard showing what share of Myntra orders ship via SJIT (FBM) versus PPMP, and how much more could move to SJIT.

## Update the dashboard

1. Download the **Seller Orders Report** from the Myntra Partner Portal (CSV).
2. Put it in `data/`.
3. Run:

```bash
pip install -r requirements.txt   # first time only
python build_dashboard.py
```

4. Open `output/myntra_sjit_dashboard.html` in a browser.

The script uses the newest file in `data/` by default.

| Command | What it does |
|---|---|
| `python build_dashboard.py` | Build from the newest report in `data/` |
| `python build_dashboard.py --combine` | Merge every report in `data/` (duplicates removed on `order line id`, newest file wins), e.g. to stitch weekly pulls into a month |
| `python build_dashboard.py --report path/to/file.csv` | Build from one specific file |
| `python build_dashboard.py --out docs/index.html` | Write the HTML somewhere else |

## What gets counted

An order line is counted only if **all** of these are blank: `cancellation reason id fk`, `cancellation reason`, `cancelled on`, `rto creation date`, `lost date`, `return creation date`. Each counted line is treated as one unit. `po_type = FBM` is SJIT; `PPMP` ships from NH warehouses.

## Config (edit these, no code changes needed)

| File | Purpose |
|---|---|
| `config/warehouse_id_mapping.xlsx` | Warehouse ID → name / state. Add rows when Myntra opens new FCs or M-Now stores. Names containing "Mnow" are treated as M-Now dark stores, names ending "- NH" as our warehouses. |
| `config/category_mapping.xlsx` | **Style ID → NH category** used for every category view. Columns: `Style ID`, `nh_sku`, `category`. When the script prints unmapped styles, add them here; until then they fall back to Myntra's article type. |
| `config/sku_mapping.json` | Secondary NH SKU → sub-category lookup, used only to help flag fresh SKUs. |
| `config/fresh_skus.json` | Fresh SKUs that always ship PPMP and are excluded from the eligible SJIT %. Matched by exact SKU, NH category (e.g. Henna, Lepa, Nutrimask, Face Mura), product-name pattern or SKU prefix. |

## Dashboard tabs

- **Overview**: SJIT % (all orders, eligible, GMV), day-on-day trend, category split, SJIT orders by Myntra location.
- **Potential**: SJIT % if every eligible SKU hit a target fill rate (slider), split into *replenish* (already in SJIT, under-stocked) and *onboard* (eligible, never shipped via SJIT).
- **Categories**: SJIT % by NH category.
- **SKU Tracker**: every SKU with its flag (Onboard / Replenish / Top up / Healthy / Low volume / Fresh).
- **Warehouses & Regions**: Myntra FCs vs M-Now, PPMP by NH warehouse, SJIT % by customer zone and state, SKU × FC heatmap.
- **Actions**: suggested SJIT quantity per SKU (DRR × days of cover × fill rate) split by North / West / South / East, exportable to CSV. No inventory file is used, so subtract current Myntra stock before consigning.

The HTML also has an **Upload a report** button to view a different report in the browser without rebuilding.

## Notes

- Charts load Chart.js and SheetJS from cdnjs, so the dashboard needs an internet connection to render.
- `data/` and `output/` are git-ignored by default because they contain order and sales data. If you publish via GitHub Pages, keep the repo private or be aware the page will be public.
