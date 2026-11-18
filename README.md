# Nat Habit × Myntra — SJIT Intelligence Dashboard

Shows what share of Myntra orders ship via SJIT, what it could be if more shelf-stable SKUs went to Myntra FCs, how each FC / FDC is being used, and what to send where.

## Daily update

1. Download from the Myntra Partner Portal:
   - **Seller Orders Report** (CSV) — every day
   - **Seller Inventory Report** (CSV) — whenever you want fresh FC stock
2. Drop both into `data/`. **Don't delete old orders reports**: the script merges all of them (duplicates removed on `order line id`, newest file wins), which is how the trend history builds up. Older reports from before you started are welcome too.
3. Double-click `update.bat` (builds `docs/index.html` and pushes it to GitHub Pages), or run by hand:

```bash
python build_dashboard.py                          # → output/myntra_sjit_dashboard.html
python build_dashboard.py --out docs/index.html    # → the GitHub Pages copy
```

The terminal prints a summary and any **styles that sold but are missing from the MPM** (also saved to `output/mpm_gaps.csv`).

## Config files

| File | What it is | When to update |
|---|---|---|
| `config/MPM_Myntra.xlsx` | MPM workbook. The script uses the Myntra sheet (Sheet2): style ID = col B, NH SKU = col C, shelf life = col J, category = col U. | When the dashboard / terminal flags MPM gaps: replace the file with the updated MPM. |
| `config/warehouse_id_mapping.xlsx` | Myntra warehouse ID → name, state. "Mnow" in the name = M-Now store, "- NH" = our warehouse. | When an unmapped warehouse ID is flagged. |
| `config/myntra_network.json` | Onboarded FCs and FDCs, their region, warehouse IDs, and the primary FC per region used in the send plan. | When you onboard a new FC/FDC or find the ID of one listed without an ID. |

## Rules used

- **Counted orders**: cancellation reason id, cancellation reason, cancelled on, RTO creation date, lost date and return creation date all blank. One order line = one unit.
- **SJIT eligibility by shelf life** (from the MPM): 12 m and 9 m = easy; 6 m = watch expiry; under 6 m = not sent to SJIT.
- **Potential**: Scenario A sends 9 & 12-month SKUs, Scenario B adds 6-month SKUs. Each eligible SKU is assumed to reach the target fill rate (slider, default 80%).
- **Send quantity per FC** = region DRR × cover − current stock in that region's FCs. Cover: 12 m → 2 months of DRR, 9 m → 1.5 months, 6 m → 1 month. DRR uses all orders from customers in the FC's region over the last 30 days (or the whole period if shorter), because each FC also feeds the M-Now stores and FDCs near it. Only FC stock counts; M-Now and FDC stock are Myntra's call.
- **Regions**: North → Gurgaon FCs (+ Lucknow FDC), West → Mumbai FC, South → Bangalore FCs (+ Chennai / Coimbatore FDCs), East → Kolkata Stellar FC.
- **Trends**: last 7 days vs the 3 weeks before (with under 3 weeks of data, the period is split in half).

## Tabs

Overview · Potential · Network (FC/FDC utilisation, region demand vs fulfilment, flow matrix, cities, M-Now, stock pivot) · Trends · Send Plan (export CSV) · SKU Tracker · Categories · MPM Gaps

## Privacy

`data/` and `output/` are git-ignored: order reports contain customer city and pincode. `docs/index.html` (the GitHub Pages copy) is pushed and is public if the repo is public.
