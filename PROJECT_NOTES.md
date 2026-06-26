# Samsung Purchases Pipeline — Project Notes / Continuation Guide

This file is for picking up this project on another machine/session. It summarizes
the current state, recent fixes, and pending tasks as of 2026-06-12.

## Project basics
- **GCP Project:** `psi-reports-493216`, **BigQuery Dataset:** `sam`
- **App repo:** https://github.com/Wajahat698/bigquery.git (branch `main`)
- **Local path:** `/Users/apple/Downloads/bigquery/app.py`
- **Venv:** `/Users/apple/Downloads/samsung_pipeline_venv`
- **Run locally:**
  ```bash
  cd ~/Downloads/bigquery && source ~/Downloads/samsung_pipeline_venv/bin/activate
  streamlit run app.py --server.port 8501 --server.headless true
  ```
- **BQ auth (local):** ADC via `gcloud auth application-default login` (wajahat698@gmail.com)
- **BQ auth (Streamlit Cloud):** `st.secrets["gcp_service_account"]` using `psi-portal-app@psi-reports-493216.iam.gserviceaccount.com`

(Full architecture / schema / table list is in the separate memory file
`project_psi_samsung_pipeline.md` — ask Claude to recall it, or see below for a copy.)

## Background: IMEI DB migration (SQL Server -> BigQuery)
Working with "Shakeel Bhai" to migrate IMEI/sales/purchase data from SQL Server
(`IMEI` database, server `10.254.10.27`) to BigQuery, eliminating Google Sheets.
**Policy: NO Google Sheet -> BigQuery pushes going forward — all data must come
from SQL Server directly.** All financial/inventory numbers must match SQL Server
exactly — always cross-verify with SQL Server (SSMS queries) before trusting a number.

## Recent fix (commit 63ba92d, 2026-06-12) — DONE & PUSHED
**Problem:** Final Summary tab's BALANCE column showed artificial negative values
(e.g. A366 8+256GB = -1065) because the rolling INHAND calc started the
purchases/sales delta range AFTER the baseline opening_balance week (W16),
but `opening_balance` represents stock at the START of W16 — so W16's own
purchases/sales must also be rolled forward.

**Fix:** in `build_final_summary()`, `delta_weeks` now includes `baseline_week`
itself:
```python
delta_weeks = ([baseline_week] + weeks_strictly_between(baseline_week, week_no)) \
    if (baseline_week is not None and baseline_week != week_no) else []
```

**Verified for Y2026W24:** A366 8+256GB went from -1065 to +365 (matches SQL
Server). Total negative-balance models dropped from 7 to 2.

### Remaining 2 negative balances — investigated, NOT bugs
- **F761 (Z Flip 7-FE) 8+256GB = -18**: 0 purchases ever recorded in
  `PURC_DATABASE` for this SKU (confirmed in both BigQuery and SQL Server),
  but 18 units sold (ACTUAL_SALE_IMEIS, matches BigQuery exactly: W16=1,
  W20=12, W22=5). **Genuine missing-purchase-data gap in SQL Server.**
- **S731 (S25 FE) 8+256GB = -345**: opening_balance(300) + purchases W17-18
  (250) = 550 available, but ACTUAL_SALE_IMEIS W17-22 = 895 sold (matches
  BigQuery exactly). Demand exceeds supply by 345 even before counting the
  `sales` table rows (which got deduped as overlaps). **Either opening_balance
  is too low, purchases under-recorded, or genuinely oversold — needs Shakeel
  Bhai to check SQL Server purchase records for this SKU.**

## Earlier in this session (also done)
- Synced `sales` and `actual_sales` BigQuery tables from fresh SQL Server CSV
  exports for W17-W24 (with backup tables `sales_backup_20260612`,
  `actual_sales_backup_20260612` created first).
- Fixed double-counting: `sales` and `actual_sales` overlap 91-99% by IMEI —
  all queries now dedupe via `QUALIFY ROW_NUMBER() OVER (PARTITION BY IMEI...) = 1`.
- Fixed `opening_balance` column containing nested dict/struct artifacts —
  unwrapped via `.apply(lambda v: v.get("opening_balance") if isinstance(v, dict) else v)`.

## Open / pending task — Daily Sales & Purchase Reports (March-May 2026)
Shakeel Bhai asked for:
1. Daily sales report, March-May 2026
2. Daily purchase report, March-May 2026

**Status: confirmed data EXISTS, report not yet built.**
- `sales` table: 141,271 rows with valid `YYYY-MM-DD` dates in this range
  (weeks Y2026W10-Y2026W22). ~24% of all-time rows have garbage `DATE` values
  (customer names/week numbers leaked in from a CSV comma-parsing issue) —
  but the March-May 2026 daily counts sampled look clean.
- `purchases` table: `PurchaseDate` column has **two mixed formats** —
  1,421,079 rows `YYYY-MM-DD` + 812,384 rows `M/D/YYYY` (+ 1,000 other).
  For March-May 2026: 108,177 rows in `YYYY-MM-DD` format + 81,612 rows in
  `M/D/YYYY` format = ~189,789 total. **Need to normalize both formats before
  building the daily report.**

**Next step:** build both reports (daily sales count/units and daily purchase
count/units, March 1 - May 31 2026, broken down by SKU/model), export as
Excel/CSV. Was about to start this when this notes file was requested.

## Conversation history
Full conversation transcripts are stored by Claude Code locally at:
`/Users/apple/.claude/projects/-Users-apple/*.jsonl`
These are NOT portable/human-readable project files — they're Claude Code's
session logs. This PROJECT_NOTES.md plus the memory file
`project_psi_samsung_pipeline.md` (in `/Users/apple/.claude/projects/-Users-apple/memory/`)
are the durable summaries meant to carry context across machines/sessions.
If you copy this whole `bigquery/` folder (with this file) to a new machine and
open a new Claude Code session there, paste this file's contents (and ask
Claude to read the memory file too) to restore full context.
