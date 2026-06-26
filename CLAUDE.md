# CLAUDE.md — Samsung Purchases Pipeline
## Reference for future Claude Code sessions

---

## Environment

| Thing | Value |
|---|---|
| GCP Project | `psi-reports-493216` |
| BigQuery Dataset | `sam` |
| Python | `/opt/anaconda3/bin/python3` |
| gcloud config dir | `/Users/wajahat/gcloud-config` |
| ADC credentials | `/Users/wajahat/gcloud-config/application_default_credentials.json` |
| Secrets (service account) | `/Users/wajahat/Downloads/bigquery 2/.streamlit/secrets.toml` |
| App repo | https://github.com/Wajahat698/bigquery (branch `main`) |
| Local path (per old docs) | `/Users/wajahat/Downloads/bigquery 2/` |

**⚠️ Machine/user note (found 2026-06-17):** On this machine the actual git
clone with the GitHub remote lives at **`/Users/apple/Downloads/bigquery/`**
(not `/Users/wajahat/...` — that path doesn't exist here). Secrets are at
`/Users/apple/Downloads/bigquery/.streamlit/secrets.toml`. The
`/Users/wajahat/gcloud-config/...` ADC credentials path still works (shared
across user dirs). There's also a near-duplicate, non-git copy at
`/Users/apple/Downloads/bigquery-main/` used for editing/reference — **always
mirror any `app.py` edit to both** `/Users/apple/Downloads/bigquery/app.py`
*and* `/Users/apple/Downloads/bigquery-main/app.py`, but only commit/push
from the `bigquery` (git) one.

**Always prefix one-off scripts with:**
```bash
GOOGLE_APPLICATION_CREDENTIALS=/Users/wajahat/gcloud-config/application_default_credentials.json \
  /opt/anaconda3/bin/python3 << 'PYEOF'
...
PYEOF
```

**Never use `.to_dataframe()`** — pyarrow/protobuf conflict in Anaconda.  
Use `list(bq.query(...))` and iterate `Row` objects directly.

**Git binary:** `/usr/bin/git` is blocked by an Xcode license agreement.  
Use `/Applications/Xcode.app/Contents/Developer/usr/bin/git` instead,  
or ask the user to run `sudo xcodebuild -license` in Terminal to fix permanently.

---

## Google Sheets — Key IDs

| Sheet | ID | Purpose |
|---|---|---|
| BACKWORKING | `1AU_lu3634-UB0vDYpBPXM7xCNW6TyjR9axcoWQ7_gKg` | Primary working sheet — source of truth, app writes here |
| SYNCED | `1XaJYcSsaqHvsQBNBCO7z0ij-tZUXOHWlsCL8WdUjzww` | Mirror sheet — auto-synced on every promo upload |
| Copy-of-BACKWORKING | `1_sWEy7jwEFpiOJrxnrgBRIC1PxePmovWidK49dw6nus` | **A manually-made COPY of BACKWORKING (titled "Copy of SAMSUNG WEEKLY - BACKWORKING - June 12, 7:08 PM"), NOT auto-synced by the app.** This is what the user shares with the **client** to review. If the client reports a data issue, check/fix this sheet too — the app's auto-sync does NOT touch it. |

All three sheets share the **same tab GIDs**:
- `183516201` → `PROMOTION DATA`
- `2031721001` → `PROMOTION - SM`

**BACKWORKING column C** (SAGE CODE) is **protected** — never write to it directly.  
**SYNCED sheet column C** is also **protected against full-row range writes** (e.g. `A1:J1` fails with
"trying to edit a protected cell"), even though CLAUDE.md historically said it wasn't. **Individual
cell writes to C work fine; only multi-column ranges that include C fail.** Always write in two
separate calls — `A:B` then `D:J` — skipping C entirely, same pattern as BACKWORKING. This applies to
ad-hoc repair scripts too, not just `sync_promo_rows_to_sheet()`.

**Verified 2026-06-22 via the Sheets API (`fetch_sheet_metadata` protectedRanges), checked on all
3 sheets (new primary, legacy BACKWORKING, SYNCED):** only `PROMOTION DATA` column C has an actual
protected range. **`PROMOTION - SM` has NO protected columns on any of the 3 sheets** — the "column
C protected" rule above only strictly applies to `PROMOTION DATA`. The app code still skips C on
both tabs (writes `A:B` then `D:J` everywhere) for consistency/safety — that's not wrong, just more
conservative than required on `PROMOTION - SM`. Don't assume protection without checking — re-verify
via `fetch_sheet_metadata()` if a sheet write fails unexpectedly or behaves differently than documented.

**Row insertion/deletion (`ws.delete_rows()` / inserting rows) is BLOCKED by sheet protection on both
BACKWORKING and SYNCED.** To "delete" bad rows in a repair script: overwrite them in place with
correct data (via `A:B` + `D:J` writes) and `clear` (write blank values to) any leftover rows if the
new data is shorter than the old. Never assume `delete_rows` will work — test on a tiny throwaway
range first.

---

## app.py — Key Constants

```python
BQ_PROJECT                  = "psi-reports-493216"                          # line ~45
BQ_DATASET                  = "sam"                                          # line ~46
BACKWORKING_SHEET_ID        = "1NkEOEjGy9czmaLOf5AgDWQSIgZUVceLxPRosGuj_8_Y"  # line ~95 — PRIMARY (changed 2026-06-22)
LEGACY_BACKWORKING_SHEET_ID = "1AU_lu3634-UB0vDYpBPXM7xCNW6TyjR9axcoWQ7_gKg"  # line ~96 — old BACKWORKING, now a mirror
SYNCED_SHEET_ID             = "1XaJYcSsaqHvsQBNBCO7z0ij-tZUXOHWlsCL8WdUjzww"  # line ~97
PROMO_SHEET_TABS = {
    "promotion_data": 183516201,   # "PROMOTION DATA" tab
    "promotion_sm":   2031721001,  # "PROMOTION - SM" tab
}                                                                     # line ~547
FAMILY_CODE_PREFIXES = { ... }                                        # line ~602
```

**⚠️ Sheet hierarchy changed 2026-06-22:** The primary working sheet (`BACKWORKING_SHEET_ID`)
is now `1NkEOEjGy9czmaLOf5AgDWQSIgZUVceLxPRosGuj_8_Y` (titled "wajahat Copy of SAMSUNG WEEKLY -
BACKWORKING" in Drive — confusing name, but it's the new primary, not a backup). It has identical
tab gids to the old BACKWORKING sheet, so all gid-based reads/writes kept working unchanged after
the swap. The **old** BACKWORKING sheet (`1AU_lu3634...`) is now `LEGACY_BACKWORKING_SHEET_ID`, a
mirror alongside `SYNCED_SHEET_ID` — both get written by `sync_promo_rows_to_sheet()` and
`push_df_to_sheet()` (Final Summary sync) after the primary write, best-effort (`try/except`,
warn-don't-fail). **The new primary sheet's formula-driven columns (stock status, claim
calculations) are broken (`#REF!`/`#N/A`)** because copying the spreadsheet broke its cross-sheet
references — this was a known/accepted tradeoff since the app only writes/reads the raw entered
columns (A:J), not the formula columns. A one-off backfill (`backfill_promo_from_new_sheet.py`)
ingested rows present in the new primary sheet but missing from BigQuery `promotion_data`/
`promotion_sm`, deduped on `(PROMOTION_NO, SAMSUNG_CODE, normalized START_DATE, normalized
QTY_ON_PROMOTION)` — date/number normalization is required because the sheet has mixed date
formats (`M/D/YYYY`, `MM-DD-YY`, etc.) that don't match BigQuery's stored `YYYY-MM-DD HH:MM:SS`.
Both BQ tables also got a new `SENT_DATE` column (nullable STRING), populated going forward in
`push_promotion_rows()` for newly-pushed promotions; backfilled historic rows have `SENT_DATE =
NULL` since their actual sent date is unknown.

---

## Promotion Sheet Column Rules

### BACKWORKING sheet (PROMOTION DATA / PROMOTION - SM)
Write **A, B, D, E, G, I, J** — skip everything else:
- **A** = Promotion title
- **B** = Promotion reference no (e.g. `A0S3A0A260000849`)
- **C** = SKIP — protected VLOOKUP (sage code auto-fills from D)
- **D** = Samsung code (real code like `SM-A075MZKDGTO`, NOT family code like `G-A07`)
- **E** = Start date (`MM-DD-YY`)
- **F** = SKIP — week auto-fills
- **G** = End date (`MM-DD-YY`)
- **H** = SKIP — week auto-fills
- **I** = QTY on promotion
- **J** = Per unit price (`$X.XX`)
- **K+** = SKIP — auto-fill (total value, claimed, remaining balance, etc.)

### SYNCED sheet (same tabs, same gids)
Same column rules. C is writable but VLOOKUP handles it from D if left blank.

---

## Family Code → Real Samsung SKUs (expand to ALL, qty SPLIT evenly — see bug history below)

PDF bulletins list a family code (e.g. `G-A07`) with **one total target qty** for the whole family.
The correct behavior (fixed 2026-06-17, see "Session 2026-06-17" below) is: expand to **every real
SKU** in that family from the MODELS sheet, and **split the total qty evenly** across them (remainder
distributed to the first SKUs). Do **NOT** write a single row with the family's first SKU and the
full total qty — that was the original bug.

| Family Code | `FAMILY_CODE_PREFIXES` value | Notes |
|---|---|---|
| G-A07 | `['A075']` | |
| G-A17 | `['A176']` per code, **but BACKWORKING's actual historical promo data for A17 uses `A175` codes** (e.g. `SM-A175FZKFGTO`), not `A176`. If a verification script checks A17 rows by `SM-A176` prefix it will wrongly report 0 SKUs / mismatch — check both, or better, just compare against whatever BACKWORKING already has for that promo. |
| G-A36-5G | `['A366']` | |
| G-A37-5G | `['A376']` | |
| G-A56-5G | `['A566']` | |
| G-A57-5G | `['A576']` | |
| G-S25-ULTRA | `['S938']` | **Exception:** in promo `A0S3A0A260000858` specifically, "G-S25-ULTRA" target qty (130) is split across **both S731 (Galaxy S25 FE) and S938 (Galaxy S25 Ultra) SKUs combined** (4 S731 rows @12 + 7 S938 rows totaling 82 = 130). Don't assume G-S25-ULTRA is always S938-only — check what BACKWORKING already contains for that specific promo before assuming a single prefix. |
| G-S26 | `['S942']` | |
| G-S26+ | `['S947']` | |
| G-S26-ULTRA | `['S948']` | |

In code: `expand_family_rows_for_sync()` in `app.py` (~line 636) handles this automatically — see
exact current logic below.

```python
def expand_family_rows_for_sync(to_insert):
    expanded = []
    for r in to_insert:
        skus = expand_family_code(r["MODEL_CODE"])
        if not skus:
            expanded.append(r)
            continue
        total_qty = int(r.get("QTY", 0))
        n = len(skus)
        base_qty = total_qty // n
        remainder = total_qty % n
        for i, sku in enumerate(skus):
            row = dict(r)
            row["MODEL_CODE"] = sku["samsung_code"]
            row["QTY"] = base_qty + (1 if i < remainder else 0)
            expanded.append(row)
    return expanded
```

Local snapshot: `models_sheet.csv` in repo root (935 rows, pulled 2026-06-16).
**CSV column indices (0-indexed, header row skipped):** `0=MARKET, 1=TYPE, 2=BRAND, 3=MODEL,
4=Sage Code, 5=DESCRIPTION, 6=SAMSUNG CODE, 7=Color/SIM`. (A previous version of this doc said
3=MODEL/4=Sage/6=SAMSUNG which is actually correct — but a quick ad-hoc script in this session
first guessed 2/3/5 and got 0 matches for every family code. Always print a header dump and verify
indices before trusting them in a throwaway script.)

---

## Promotion Upload Flow (Streamlit UI)

1. User uploads PDF bulletin → `parse_promotion_pdf()` extracts rows
2. `check_promotion_rows()` compares vs BQ — finds duplicates / new versions
3. **"Push to BigQuery" button** → `push_promotion_rows()` inserts to `promotion_data` or `promotion_sm`
4. After push, UI shows: extension detections (blue), price changes (salmon), SKU coverage gaps
5. **"Sync these rows to the BACKWORKING sheet" button** → `sync_promo_rows_to_sheet()`:
   - Writes to **BACKWORKING** sheet (primary)
   - **Also writes to SYNCED sheet automatically** (added 2026-06-16, `app.py` line ~771)
   - Family codes resolved to first real Samsung code before writing
   - `time.sleep(2)` between every batch write (rate limit)

**Note:** The button label still says "BACKWORKING sheet" but it now writes to both. Could update label in future.

**Which BQ table?**
- `promotion_data` = Regular / Sell-Thru market promotions
- `promotion_sm` = Special Market (detected from bulletin title containing "Special Market")

---

## Current Sheet State (as of 2026-06-17, post-fix — all verified ✅)

PROMOTION DATA tab (gid=183516201), for the 5 affected promos, on **both** SYNCED and
Copy-of-BACKWORKING:

| Promo | Rows | Per-family qty split (verified totals match bulletin) |
|---|---|---|
| A0S3A0A260000755 | 45 | A36-5G=430, A37-5G=415, A56-5G=860, A57-5G=1210, S26=65, S26+=5, S26-ULTRA=1300 |
| A0S3A0A260000814 | 15 | A07=4860, A17=2970 |
| A0S3A0A260000836 | 6 | S25-ULTRA=880 |
| A0S3A0A260000849 | 9 | A07=520, A17=426 |
| A0S3A0A260000858 | 38 | A37-5G=70, A57-5G=80, S25-ULTRA(S731+S938)=130, S26=55, S26+=50, S26-ULTRA=95 |

SYNCED PROMOTION DATA total: **3486 rows**. Copy-of-BACKWORKING PROMOTION DATA total: **3457 rows**.
(Row counts differ between the two sheets because Copy-of-BACKWORKING doesn't have the unrelated
promo `A0S3A0A260000823` block that SYNCED keeps in the middle of this row range — don't assume
the two sheets have identical row numbers for the same promo.)

PROMOTION-SM tab (gid=2031721001) on SYNCED still has leftover bad rows from the pre-fix session
for promo `A0S3A0A260000807` and possibly stray promotion_data rows mixed in — **not fully
re-verified after the 2026-06-17 fix**. If the client/user reports SM-tab issues, re-run the same
expand-and-split-qty repair there.

---

## Promotions Done in Session (2026-06-16) — All 6 Verified Complete (BQ push)

| Ref No | Title | BQ Table | BACKWORKING | SYNCED |
|---|---|---|---|---|
| A0S3A0A260000755 | W22-W24 PLANET OM Sell-Thru | promotion_data | ✅ | ✅ |
| A0S3A0A260000807 | PLANET W24 AR Special Market | promotion_sm | ✅ | ✅ |
| A0S3A0A260000814 | W24 A17/A07 Sell-Thru | promotion_data | ✅ | ✅ |
| A0S3A0A260000836 | PLANET W24 S25 Ultra Sell-Thru | promotion_data | ✅ | ✅ |
| A0S3A0A260000849 | PLANET W25 A17/A07 Sell-Thru | promotion_data | ✅ | ✅ |
| A0S3A0A260000858 | W25-W27 PLANET OM Sell-Thru | promotion_data | ✅ | ✅ |

---

## Code Changes Made — Session 2026-06-16

### Commit `f4d559c` — Auto-sync to SYNCED sheet
- Added `SYNCED_SHEET_ID = "1XaJYcSsaqHvsQBNBCO7z0ij-tZUXOHWlsCL8WdUjzww"` at `app.py:96`
- Extended `sync_promo_rows_to_sheet()` (`app.py:771–784`) to also write to SYNCED sheet after every BACKWORKING write
- Both writes use same column layout (A:B then D:J), rate-limited with `time.sleep(2)`

### Commit `e413289` — models_sheet.csv
- Added `models_sheet.csv` to repo root — local copy of MODELS tab (935 rows)
- Use for SKU lookups in one-off scripts without hitting Sheets API

---

## Code Changes Made — Session 2026-06-17 — **Family-code expansion bug, fully fixed**

**The bug (client-reported, in Hindi/Urdu: "model k according promo ban rahi hai, sku k according
ne" = "promo is being built per-model, not per-SKU"):** `expand_family_rows_for_sync()` originally
took only `skus[0]` (the first SKU in the family) and wrote the **entire family's total qty** onto
that single row. So `G-A07` qty=520 in the bulletin became ONE row (`SM-A075MZKDGTO`, qty=520)
instead of 6 rows (one per real A07 SKU) with the qty split across them.

**The fix** (commit `11fd0b3` on `https://github.com/Wajahat698/bigquery`, also mirrored to
`bigquery` and `bigquery-main` local copies): `expand_family_rows_for_sync()` now expands to **every**
real SKU for the family (via `expand_family_code()`) and **splits qty evenly** (floor division,
remainder distributed to the first N rows) — see code block above in the Family Code section.

**Mid-session false alarm:** at one point the user said "don't split, give each SKU the full total
qty" — this was wrong/a miscommunication; a screenshot of the *actual correct* BACKWORKING data
later confirmed **split is correct** (qty=87 per A07 SKU for promo 849, not 520 per SKU). Don't
re-introduce "full qty per SKU" — it was tried and explicitly reverted.

**Sheets repaired manually (one-off Python scripts, not via the Streamlit UI)** for the 5 affected
promos (755, 814, 836, 849, 858) on:
1. SYNCED (`1XaJYcSsaqHvsQBNBCO7z0ij-tZUXOHWlsCL8WdUjzww`) — done, verified ✅
2. Copy-of-BACKWORKING (`1_sWEy7jwEFpiOJrxnrgBRIC1PxePmovWidK49dw6nus`, the client-facing copy) — done, verified ✅

The original BACKWORKING sheet (`1AU_lu3634...`) already had the correct split-qty data from an
earlier (pre-bug) session and was used as the **source of truth** to repair the other two sheets —
it was NOT modified in this session.

**Verification method:** for each promo, summed qty across all SKUs sharing the family's Samsung
prefix and compared to the bulletin's total target. All matched exactly (see table above). A
handful of SKUs in promo 755's A366 family have blank qty — pre-existing in BACKWORKING, not
introduced by this fix.

---

## Gotchas — Do Not Repeat

| Problem | Fix |
|---|---|
| Family codes like `G-A07` written as Samsung code | Always call `expand_family_rows_for_sync()` before writing — expands to ALL real SKUs in the family, splitting qty evenly (not just the first SKU with full qty — that was the 2026-06-17 bug) |
| Rate-limit 429 from Sheets API | `time.sleep(2)` between every `ws.update()` call — never skip |
| BACKWORKING col C protected | Skip C entirely (VLOOKUP fills it from D) |
| **SYNCED col C is ALSO protected for multi-column range writes** | `A1:J1`-style writes that include col C fail with "protected cell" error, even though older docs said C was writable there. Write `A:B` then `D:J` as two separate calls — same as BACKWORKING. Single-cell writes to C alone work, but never include C in a wider range write. |
| **Row delete/insert blocked by protection on BACKWORKING & SYNCED** | `ws.delete_rows()` fails with "protected cell or object" on both sheets. To fix bad rows: overwrite in place (`A:B` + `D:J` writes) with correct data, then `clear` (write blanks to) any leftover rows if new data is shorter. Test on a throwaway 1-2 row range before committing to a big repair script. |
| `ws.worksheet("title")` can break on rename | Use `ws.get_worksheet_by_id(gid)` — gids are stable |
| `.to_dataframe()` crashes Anaconda | pyarrow/protobuf conflict — use `list(bq.query(...))` and iterate rows |
| `/usr/bin/git` exits 69 (Xcode license) | Use `/Applications/Xcode.app/Contents/Developer/usr/bin/git` |
| MODELS sheet / `models_sheet.csv` columns | 0=MARKET, 1=TYPE, 2=BRAND, **3=MODEL, 4=Sage Code**, 5=DESCRIPTION, **6=SAMSUNG CODE**, 7=Color/SIM. Verify with a header dump before trusting — guessing wrong indices silently returns 0 SKUs for every family (looks like "no bug" but is actually "script is broken"). |
| There's a 3rd sheet besides BACKWORKING/SYNCED | "Copy-of-BACKWORKING" (`1_sWEy7jwEFpiOJrxnrgBRIC1PxePmovWidK49dw6nus`) is what gets shown to the **client** and is NOT touched by the app's auto-sync. If client reports an issue, check this sheet specifically — fixing BACKWORKING/SYNCED alone won't be visible to them. |
| G-A17 family prefix mismatch | Code's `FAMILY_CODE_PREFIXES` says `A176`, but real historical promo data in BACKWORKING for A17 uses `A175` codes. Don't assume the code constant always matches what's actually in the sheets — cross-check actual sheet data when verifying. |
| Local repo confusion | This machine has the real git clone at `/Users/apple/Downloads/bigquery/` (NOT the `/Users/wajahat/...` path from old docs) plus a non-git mirror at `/Users/apple/Downloads/bigquery-main/`. Edit+commit from `bigquery/`, but keep `bigquery-main/app.py` in sync too since it may be referenced/edited directly in some sessions. |

---

## gspread Pattern for Sheet Writes (rate-limit safe)

```python
import toml, gspread, time
secrets = toml.load('/Users/wajahat/Downloads/bigquery 2/.streamlit/secrets.toml')
gc = gspread.service_account_from_dict(dict(secrets['gcp_service_account']),
    scopes=['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'])
sh = gc.open_by_key(SHEET_ID)
ws = sh.get_worksheet_by_id(GID)   # always use gid, not tab name

start = len(ws.get_all_values()) + 1
end   = start + len(rows) - 1

ws.update(values=ab_rows, range_name=f'A{start}:B{end}', value_input_option='RAW')
time.sleep(2)
ws.update(values=dj_rows, range_name=f'D{start}:J{end}', value_input_option='RAW')
time.sleep(2)
```

## BQ Query Pattern (no pyarrow)

```python
from google.cloud import bigquery
bq = bigquery.Client(project='psi-reports-493216')
rows = list(bq.query("""
    SELECT PROMOTION_NO, SAMSUNG_CODE, START_DATE, END_DATE, QTY_ON_PROMOTION, PER_UNIT
    FROM `psi-reports-493216.sam.promotion_data`
    WHERE PROMOTION_NO = 'A0S3A0A260000849'
"""))
for r in rows:
    print(r.PROMOTION_NO, r.SAMSUNG_CODE, float(r.PER_UNIT))
```
