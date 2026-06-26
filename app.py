import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import openpyxl
import pdfplumber
import io
import re
import json
import uuid
from datetime import datetime, date, timezone

st.set_page_config(page_title="Samsung Purchases Pipeline", page_icon="📦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #080810; color: #e0e0f0; }

/* Hide Streamlit Watermarks and Menus */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.viewerBadge_container__1QSob {display: none !important;}
.styles_viewerBadge__1yB5_ {display: none !important;}
.stApp [data-testid="stToolbar"] {display: none;}
h1, h2, h3 { font-family: 'IBM Plex Mono', monospace !important; color: #00e5ff !important; }
.block-container { padding-top: 1.5rem; }
.step-card { background: #0f0f1c; border: 1px solid #1a1a30; border-radius: 8px; padding: 1.2rem 1.5rem; margin-bottom: 0.8rem; }
.step-card.success { border-left: 3px solid #00e676; }
.step-card.error   { border-left: 3px solid #ff1744; }
.step-card.warning { border-left: 3px solid #ffab00; }
.step-card.idle    { border-left: 3px solid #2a2a45; }
.step-title { font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem; color: #888; letter-spacing: 0.1em; text-transform: uppercase; }
.step-body  { font-size: 0.95rem; color: #e0e0f0; margin-top: 0.1rem; }
.step-meta  { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #555; margin-top: 0.2rem; }
.metric-row { display: flex; gap: 1.5rem; margin: 1rem 0; flex-wrap: wrap; }
.metric-box { background: #0f0f1c; border: 1px solid #1a1a30; border-radius: 6px; padding: 0.8rem 1.2rem; min-width: 120px; text-align: center; }
.metric-val { font-family: 'IBM Plex Mono', monospace; font-size: 1.6rem; color: #00e5ff; }
.metric-lbl { font-size: 0.7rem; color: #666; text-transform: uppercase; letter-spacing: 0.1em; margin-top: 0.2rem; }
.stButton > button { background: #00e5ff !important; color: #080810 !important; font-family: 'IBM Plex Mono', monospace !important; font-weight: 600 !important; border: none !important; border-radius: 4px !important; padding: 0.6rem 2rem !important; }
.stButton > button:hover { background: #33ecff !important; }
.stButton > button:disabled { background: #1a1a30 !important; color: #444 !important; }
[data-testid="stFileUploader"] { border: 1px solid #1a1a30; border-radius: 8px; padding: 0.8rem; background: #0f0f1c; }
.stTextInput > div > div, .stSelectbox > div { background: #0f0f1c !important; border: 1px solid #1a1a30 !important; color: #e0e0f0 !important; border-radius: 4px !important; }
hr { border-color: #1a1a30 !important; }
[data-testid="stDataFrame"] { border: 1px solid #1a1a30; border-radius: 6px; }
.stTabs [data-baseweb="tab"] { color: #888; font-family: 'IBM Plex Mono', monospace; }
.stTabs [aria-selected="true"] { color: #00e5ff !important; }
</style>
""", unsafe_allow_html=True)

BQ_PROJECT  = "psi-reports-493216"
BQ_DATASET  = "sam"
TBL_STAGING = "current_week_sales"
TBL_MASTER  = "purchases"
TBL_IMEIS   = "purchase_imeis"
BQ_SCHEMA   = ["SO_NO","BILLING_NO","MATL","EIN","DUAL_IMEI_NO","PO_NO","WEEK_NO","PurchaseDate"]

for key in ["df_out","step1_done","step2_done","step3_done","step4_done","step5_done",
            "step1_msg","step2_msg","step3_msg","step4_msg","step5_msg",
            "step1_status","step2_status","step3_status","step4_status","step5_status",
            "duplicates_df","total_rows","matched_po","week_no_used","purchase_date_used"]:
    if key not in st.session_state:
        st.session_state[key] = None

def reset_pipeline():
    for key in ["df_out","step1_done","step2_done","step3_done","step4_done","step5_done",
                "step1_msg","step2_msg","step3_msg","step4_msg","step5_msg",
                "step1_status","step2_status","step3_status","step4_status","step5_status",
                "duplicates_df","total_rows","matched_po","week_no_used","purchase_date_used"]:
        st.session_state[key] = None

def step_card(num, title, status, body, meta=""):
    icons = {"success":"✅","error":"❌","warning":"⚠️","idle":"⏳","running":"🔄"}
    icon = icons.get(status, "⏳")
    meta_html = f"<div class='step-meta'>{meta}</div>" if meta else ""
    st.markdown(
        f"""<div class="step-card {status}"><div class="step-icon" style="font-size:1.4rem;min-width:2rem;text-align:center;display:inline-block">{icon}</div>&nbsp;&nbsp;<div style="display:inline-block"><div class="step-title">Step {num} — {title}</div><div class="step-body">{body}</div>{meta_html}</div></div>""",
        unsafe_allow_html=True)

def safe_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

@st.cache_resource
def get_bq_client():
    from google.cloud import bigquery
    try:
        has_secret = "gcp_service_account" in st.secrets
    except Exception:
        has_secret = False
    if has_secret:
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(st.secrets["gcp_service_account"])
        return bigquery.Client(project=BQ_PROJECT, credentials=creds)
    return bigquery.Client(project=BQ_PROJECT)

# BACKWORKING_SHEET_ID        = "1DKnokosWYv4qsMSG4pspz7JxGxfjzh8dAaEvVAgmV6w"  # primary working sheet (clean, set 2026-06-22)
BACKWORKING_SHEET_ID        = "1gHeuIVhf64PTMEU_fE0puZTJR9ejZ6cQ2gZpSJx2cRo"
# LEGACY_BACKWORKING_SHEET_ID = "1AU_lu3634-UB0vDYpBPXM7xCNW6TyjR9axcoWQ7_gKg"  # old BACKWORKING, now a mirror
LEGACY_BACKWORKING_SHEET_ID = None
# SYNCED_SHEET_ID             = "1XaJYcSsaqHvsQBNBCO7z0ij-tZUXOHWlsCL8WdUjzww"
SYNCED_SHEET_ID             = None
# NOTE: the previous primary "1NkEOEjGy9czmaLOf5AgDWQSIgZUVceLxPRosGuj_8_Y" ("wajahat Copy")
# accumulated bad rows from sync_promotion_table_append bugs and has broken cross-sheet
# formulas (#REF!/#N/A) — intentionally dropped from all sync targets. Do not re-add it.

SM_SYNC_TAB = "SM - FINAL SUMMARY (App Sync)"
RM_SYNC_TAB = "FINAL SUMMARY (App Sync)"

# Maps SM Final Summary tab column names (as seen in the sheet) back to the
# underlying "SM ..." columns in df_summary.
SM_RENAME_MAP = {
    "SM PROMO NUMBER": "PROMO NUMBER",
    "SM PROMO STATUS": "PROMO STATUS",
    "SM PROMO END DATE": "PROMO END DATE",
    "SM QTY ON PROMOTION": "QTY ON PROMOTION",
    "SM REMAINING BAL QTY": "REMAINING BAL QTY",
    "SM PROMO PER UNIT": "VALUE P/U",
    "SM CLAIMED QTY": "CLAIMED QTY",
    "SM CLAIMED VALUE": "CLAIMED VALUE",
}
SM_RENAME_MAP_INV = {v: k for k, v in SM_RENAME_MAP.items()}

@st.cache_resource
def get_gsheet_client():
    import gspread
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            scopes = ["https://www.googleapis.com/auth/spreadsheets",
                      "https://www.googleapis.com/auth/drive"]
            return gspread.service_account_from_dict(creds_dict, scopes=scopes)
    except Exception:
        pass

    # Fallback to ADC if no secrets.toml is present
    try:
        import google.auth
        scopes = ["https://www.googleapis.com/auth/spreadsheets",
                  "https://www.googleapis.com/auth/drive"]
        credentials, _ = google.auth.default(scopes=scopes)
        return gspread.authorize(credentials)
    except Exception:
        return None

def _clear_and_write_tab(sh, tab_title, values):
    """Overwrite (or create) a tab in spreadsheet `sh` with `values` (header + rows)."""
    try:
        ws = sh.worksheet(tab_title)
        ws.clear()
    except Exception:
        ws = sh.add_worksheet(title=tab_title, rows=str(len(values) + 10), cols=str(len(values[0]) + 2))
    ws.update(values, value_input_option="RAW")

def push_df_to_sheet(df, tab_title):
    """Overwrite (or create) a tab in the BACKWORKING sheet with df's contents,
    then mirror the same write to LEGACY_BACKWORKING_SHEET_ID and SYNCED_SHEET_ID
    (best-effort — a mirror failure is warned, not fatal)."""
    gc = get_gsheet_client()
    if gc is None:
        raise RuntimeError("Google Sheets credentials not configured (gcp_service_account missing).")
    import math
    clean = df.astype(object).where(pd.notna(df), "")
    clean = clean.map(lambda v: "" if isinstance(v, float) and math.isinf(v) else v)
    values = [list(df.columns)] + clean.values.tolist()

    sh = gc.open_by_key(BACKWORKING_SHEET_ID)
    _clear_and_write_tab(sh, tab_title, values)

    import time as _time
    for mirror_id in (LEGACY_BACKWORKING_SHEET_ID, SYNCED_SHEET_ID):
        if not mirror_id:
            continue
        try:
            _time.sleep(2)
            mirror_sh = gc.open_by_key(mirror_id)
            _clear_and_write_tab(mirror_sh, tab_title, values)
        except Exception as _e:
            st.warning(f"Mirror sync to {mirror_id} failed for tab '{tab_title}': {_e}")

# Promotion working sheets: wide, multi-block layout with a 2-row header.
# Columns C (SAGE CODE) and K-onward (claims-tracking formulas, currently #REF!/#N/A
# on the primary sheet after a broken spreadsheet copy) must NEVER be read from or
# written to by app code — only A,B,D,E,G,I,J are ever touched, same as
# sync_promo_rows_to_sheet(). This means we can't dedupe on the old UNIQUE_ID column
# (it lived at column AA, well past J) — dedupe instead on a composite key read from
# the safe columns themselves: (PROMOTION_NO, SAMSUNG_CODE, START_DATE, QTY_ON_PROMOTION).

def _blank(v):
    """A pandas NULL from a STRING BQ column can come back as None OR as a raw
    float('nan') (BigQuery Storage API behavior) — `v or ""` alone doesn't catch the
    float case since NaN is truthy in Python, which left raw NaN floats in row data
    and crashed JSON serialization ("Out of range float values are not JSON compliant").
    Always pass values through this before using them in a sheet write."""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return v

def _norm_sheet_date(v):
    """Normalize a date string (sheet 'M/D/YYYY' or BigQuery 'YYYY-MM-DD HH:MM:SS') to YYYY-MM-DD."""
    s = str(_blank(v) or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%y", "%m-%d-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", s)
    if m:
        mo, d, y = m.groups()
        y = int(y)
        if y < 100:
            y += 2000
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return s

def _norm_sheet_num(v):
    s = str(v or "").strip().replace("$", "").replace(",", "")
    if not s:
        return ""
    try:
        return str(float(s))
    except ValueError:
        return s

def sync_promotion_table_append(bq_table, sheet_tab):
    """Append BigQuery rows (from `bq_table`) that aren't yet present in `sheet_tab`,
    matching on (PROMOTION_NO, SAMSUNG_CODE, START_DATE, QTY_ON_PROMOTION) read from the
    sheet's own A/B/D/E/I columns. Only ever writes A:B and D:J — never reads or writes
    column C or anything K-onward. Existing rows/formulas are left untouched.

    Column A must be BigQuery's MODEL column (the actual promo title) — NOT MODEL_2
    (a generic per-row category like "SMART PHONE" from the PDF table header). Family
    codes (e.g. G-A07) in SAMSUNG_CODE must be expanded to real SKUs via
    expand_family_rows_for_sync() BEFORE deduping/writing — dedup matches against the
    sheet's real-SKU rows, so checking the unexpanded family code would never match an
    already-synced promo and would re-append it every run. (Both of these were bugs in
    an earlier version of this function — see project memory before changing this again.)"""
    gc = get_gsheet_client()
    if gc is None:
        raise RuntimeError("Google Sheets credentials not configured (gcp_service_account missing).")
    sh = gc.open_by_key(BACKWORKING_SHEET_ID)
    ws = sh.worksheet(sheet_tab)
    existing = ws.get_all_values()

    def existing_key(row):
        if len(row) <= 8:
            return None
        return (row[1].strip(), row[3].strip(), _norm_sheet_date(row[4]), _norm_sheet_num(row[8]))

    existing_keys = {k for k in (existing_key(r) for r in existing[2:]) if k}

    client = get_bq_client()
    df = client.query(f"""
        SELECT MODEL AS TITLE, PROMOTION_NO, SAMSUNG_CODE, START_DATE, END_DATE,
               QTY_ON_PROMOTION, PER_UNIT
        FROM `{BQ_PROJECT}.{BQ_DATASET}.{bq_table}`
        ORDER BY PROMOTION_NO, SAMSUNG_CODE
    """).to_dataframe()

    candidates = []
    for _, r in df.iterrows():
        promo_no = str(_blank(r["PROMOTION_NO"]) or "").strip()
        samsung_code = str(_blank(r["SAMSUNG_CODE"]) or "").strip()
        title = str(_blank(r["TITLE"]) or "").strip()
        if not promo_no or not samsung_code:
            continue
        qty_num = to_num(_blank(r["QTY_ON_PROMOTION"]))
        candidates.append({
            "MODEL_CODE": samsung_code,
            "QTY": int(qty_num) if qty_num is not None else 0,
            "TITLE": title,
            "PROMOTION_NO": promo_no,
            "START_DATE": r["START_DATE"],
            "END_DATE": r["END_DATE"],
            "PER_UNIT": to_num(r["PER_UNIT"]),
        })

    # WEEK columns (F, H) are never written at all — not even a blank value — so the
    # sheet's own auto-fill formula for WEEK is the only thing that ever touches them.
    ab_rows, de_rows, g_rows, ij_rows = [], [], [], []
    for r in expand_family_rows_for_sync(candidates):
        samsung_code = r["MODEL_CODE"]
        key = (r["PROMOTION_NO"], samsung_code, _norm_sheet_date(r["START_DATE"]), _norm_sheet_num(r["QTY"]))
        if key in existing_keys:
            continue
        ab_rows.append([r["TITLE"], r["PROMOTION_NO"]])
        de_rows.append([samsung_code, _norm_sheet_date(r["START_DATE"])])
        g_rows.append([_norm_sheet_date(r["END_DATE"])])
        ij_rows.append([r["QTY"], f"${r['PER_UNIT']:.2f}" if r["PER_UNIT"] is not None else ""])
        existing_keys.add(key)

    skipped = 0
    if ab_rows:
        # Find the true last row by looking at Column A (TITLE)
        last_filled_row = 0
        for i, row in enumerate(existing):
            if len(row) > 0 and str(row[0]).strip() != "":
                last_filled_row = i + 1
        
        start_row = max(last_filled_row + 1, 3) # Data always starts at row 3
        required_rows = start_row + len(ab_rows) - 1
        
        # Expand the grid if needed (RAW tabs are unprotected, so this succeeds)
        if required_rows > ws.row_count:
            ws.add_rows(required_rows - ws.row_count)
            
        end_row = start_row + len(ab_rows) - 1
        ws.update(f"A{start_row}:B{end_row}", ab_rows, value_input_option="RAW")
        ws.update(f"D{start_row}:E{end_row}", de_rows, value_input_option="RAW")
        ws.update(f"G{start_row}:G{end_row}", g_rows, value_input_option="RAW")
        ws.update(f"I{start_row}:J{end_row}", ij_rows, value_input_option="RAW")
    return len(ab_rows), skipped

def week_to_num(w):
    m = re.match(r'Y(\d{4})W(\d+)', str(w))
    if m:
        return int(m.group(1)) * 100 + int(m.group(2))
    return 0

def to_num(v):
    """Coerce a value (possibly stray text from legacy promo uploads) to a
    float, or None if it isn't numeric."""
    try:
        n = float(v)
        return None if pd.isna(n) else n
    except (TypeError, ValueError):
        return None

def load_data_sheet_from_file(file):
    """Load data sheet from uploaded Excel file.
    Expected columns: ODO# (or SELA_DO), PO # (or PO_NO), WEEK_NO, PurchaseDate (or DATE)
    """
    wb = openpyxl.load_workbook(file, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}, None, None
    headers = [str(h).strip() if h else "" for h in rows[0]]

    # Find column indices flexibly
    def find_col(names):
        for n in names:
            for i, h in enumerate(headers):
                if n.lower() in h.lower():
                    return i
        return None

    odo_idx  = find_col(["ODO#","ODO","SELA_DO","DO#"])
    po_idx   = find_col(["PO #","PO#","PO_NO","PO"])
    week_idx = find_col(["WEEK_NO","WEEK NO","WEEK"])
    date_idx = find_col(["PurchaseDate","PURCHASE DATE","DATE"])

    po_lookup   = {}
    week_no     = None
    purchase_dt = None
    week_nums   = []

    for row in rows[1:]:
        if odo_idx is not None and po_idx is not None:
            odo = row[odo_idx]
            po  = row[po_idx]
            if odo and po:
                po_lookup[str(odo)] = str(po)

        if week_idx is not None and row[week_idx]:
            w = str(row[week_idx]).strip()
            if re.match(r'Y\d{4}W\d+', w):
                week_nums.append(w)

        if date_idx is not None and row[date_idx] and purchase_dt is None:
            dt = row[date_idx]
            if isinstance(dt, (datetime, date)):
                purchase_dt = dt.strftime("%Y-%m-%d")
            else:
                try:
                    from openpyxl.utils.datetime import from_excel
                    purchase_dt = from_excel(int(dt)).strftime("%Y-%m-%d")
                except Exception:
                    purchase_dt = str(dt)

    # Pick the most recent week from the sheet
    if week_nums:
        week_no = max(week_nums, key=week_to_num)

    return po_lookup, week_no, purchase_dt

def load_data_sheet_from_url(url):
    """Load data sheet from Google Sheets public URL."""
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m:
        raise ValueError("Invalid Google Sheets URL")
    sheet_id = m.group(1)
    gid_m = re.search(r'gid=(\d+)', url)
    gid = gid_m.group(1) if gid_m else "0"
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(csv_url, dtype=str)
    headers = list(df.columns)

    def find_col(names):
        for n in names:
            for h in headers:
                if n.lower() in h.lower():
                    return h
        return None

    odo_col  = find_col(["ODO#","ODO","SELA_DO","DO#"])
    po_col   = find_col(["PO #","PO#","PO_NO","PO"])
    week_col = find_col(["WEEK_NO","WEEK NO","WEEK"])
    date_col = find_col(["PurchaseDate","PURCHASE DATE","DATE"])

    po_lookup = {}
    if odo_col and po_col:
        for _, row in df.iterrows():
            if pd.notna(row.get(odo_col)) and pd.notna(row.get(po_col)):
                po_lookup[str(row[odo_col]).strip()] = str(row[po_col]).strip()

    week_no = None
    if week_col:
        weeks = [w for w in df[week_col].dropna() if re.match(r'Y\d{4}W\d+', str(w).strip())]
        if weeks:
            week_no = max(weeks, key=week_to_num)

    purchase_dt = None
    if date_col:
        for v in df[date_col].dropna():
            purchase_dt = str(v).strip()
            break

    return po_lookup, week_no, purchase_dt

def transform(df, po_lookup, week_no, purchase_date):
    df = df.copy()
    df["PO_NO"]        = df["SELA_DO"].map(po_lookup)
    df["WEEK_NO"]      = week_no or ""
    df["PurchaseDate"] = purchase_date or ""
    return df[BQ_SCHEMA].copy()

def promo_week_str(d):
    iso = d.isocalendar()
    return f"Y{iso[0]}W{iso[1]:02d}"

# Matches a data row like: "SMART PHONE G-A36-5G June.08.2026 June.14.2026 20.00 100"
PROMO_LINE_RE = re.compile(
    r"^(.+?)\s+(\S+)\s+([A-Za-z]+\.\d{1,2}\.\d{4})\s+([A-Za-z]+\.\d{1,2}\.\d{4})\s+([\d.]+)\s+(\d+)\s*$"
)
# Matches a delta-update line like: "G-A36-5G = 100"
DELTA_LINE_RE = re.compile(r"^([A-Za-z0-9][\w\-\+/]*)\s*=\s*(\d+)\s*$", re.MULTILINE)

def _parse_promo_date(s):
    mon, day, year = s.split(".")
    return datetime.strptime(f"{mon[:3]}.{day}.{year}", "%b.%d.%Y")

def parse_promotion_pdf(file):
    """Parse a Samsung promotion bulletin PDF.

    Returns (reference_no, title, target_table, rows, delta_updates, sent_date).
    `rows` is a list of full promotion rows (PRODUCT/MODEL_CODE/dates/PER_UNIT/QTY).
    `delta_updates` is a {MODEL_CODE: new_qty} dict for "Updated Target Units" style
    bulletins that only list new quantities for an existing promotion (rows is empty
    in that case, since dates/per-unit must come from the existing BQ rows).
    `sent_date` is the bulletin's own "Date Sent :" field (datetime, or None if absent).
    """
    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        tables = page.extract_tables()

    ref_match   = re.search(r"Reference Number\s*:?\s*:?\s*(\S+)", text)
    title_match = re.search(r"Promotion Title\s*:?\s*:?\s*(.+?)(?:\s*USD)?\n", text)
    sent_match  = re.search(r"Date Sent\s*:?\s*:?\s*([A-Za-z]+\.\d{1,2}\.\d{4})", text)
    reference_no = ref_match.group(1) if ref_match else None
    title = title_match.group(1).strip() if title_match else None
    sent_date = _parse_promo_date(sent_match.group(1)) if sent_match else None

    rows = []

    # 1. Table-based extraction
    for t in tables:
        for r in t:
            if r and len(r) >= 6 and r[0] and r[1] and re.match(r"^[A-Za-z]+\.\d{1,2}\.\d{4}$", str(r[2] or "")):
                product, model, start, end, per_unit, qty = r[:6]
                try:
                    rows.append({
                        "PRODUCT": product,
                        "MODEL_CODE": model,
                        "START_DATE": _parse_promo_date(start),
                        "END_DATE": _parse_promo_date(end),
                        "PER_UNIT": float(per_unit),
                        "QTY": int(qty),
                    })
                except (ValueError, TypeError):
                    continue

    # 2. Plain-text row fallback (no table structure detected by pdfplumber)
    if not rows:
        for line in text.splitlines():
            m = PROMO_LINE_RE.match(line.strip())
            if m:
                product, model, start, end, per_unit, qty = m.groups()
                try:
                    rows.append({
                        "PRODUCT": product.strip(),
                        "MODEL_CODE": model,
                        "START_DATE": _parse_promo_date(start),
                        "END_DATE": _parse_promo_date(end),
                        "PER_UNIT": float(per_unit),
                        "QTY": int(qty),
                    })
                except (ValueError, TypeError):
                    continue

    is_sm = "special market" in (title or "").lower()
    target_table = "promotion_sm" if is_sm else "promotion_data"

    # 3. Delta-update fallback ("Updated Target Units: MODEL = qty")
    delta_updates = None
    if not rows:
        matches = DELTA_LINE_RE.findall(text)
        if matches:
            delta_updates = {code: int(qty) for code, qty in matches}

    if not reference_no:
        raise ValueError("Could not find a Reference Number in this PDF.")
    if not rows and not delta_updates:
        raise ValueError("Could not extract any promotion data from this PDF (no table rows or update list found).")

    return reference_no, title, target_table, rows, delta_updates, sent_date

def resolve_delta_updates(reference_no, delta_updates):
    """For 'Updated Target Units' bulletins: find the existing rows for these model
    codes under this PROMOTION_NO (in either table) and build full rows with the new qty,
    reusing the existing product/dates/per-unit values."""
    from google.cloud import bigquery
    client = get_bq_client()

    for target_table in ["promotion_data", "promotion_sm"]:
        table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{target_table}"
        q = f"""
            SELECT SAMSUNG_CODE, MODEL_2, START_DATE, END_DATE, PER_UNIT
            FROM `{table_id}` WHERE PROMOTION_NO = @ref
        """
        job = client.query(q, job_config=bigquery.QueryJobConfig(
            query_parameters=[bigquery.ScalarQueryParameter("ref", "STRING", reference_no)]
        ))
        existing = {}
        for r in job:
            existing.setdefault(r.SAMSUNG_CODE, dict(r))

        matched_codes = [c for c in delta_updates if c in existing]
        if not matched_codes:
            continue

        rows = []
        for code in matched_codes:
            e = existing[code]
            rows.append({
                "PRODUCT": e["MODEL_2"] or "",
                "MODEL_CODE": code,
                "START_DATE": pd.to_datetime(e["START_DATE"]).to_pydatetime(),
                "END_DATE": pd.to_datetime(e["END_DATE"]).to_pydatetime(),
                "PER_UNIT": float(e["PER_UNIT"]),
                "QTY": delta_updates[code],
            })
        unmatched = set(delta_updates) - set(matched_codes)
        return target_table, rows, unmatched

    return None, [], set(delta_updates)

def check_promotion_rows(reference_no, target_table, rows):
    """Classify PDF rows vs existing BQ rows for this promotion:
    - duplicate: an identical row (same code + qty) already exists -> skip
    - new_version: same model code exists but with different qty/values -> append as a new row (history kept)
    - new: model code not seen before for this promotion -> append
    """
    from google.cloud import bigquery
    client = get_bq_client()
    table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{target_table}"

    existing_q = f"""
        SELECT SAMSUNG_CODE, START_DATE, END_DATE, QTY_ON_PROMOTION, PER_UNIT, TOTAL_VALUE
        FROM `{table_id}` WHERE PROMOTION_NO = @ref
    """
    job = client.query(existing_q, job_config=bigquery.QueryJobConfig(
        query_parameters=[bigquery.ScalarQueryParameter("ref", "STRING", reference_no)]
    ))
    existing_by_code = {}
    for r in job:
        existing_by_code.setdefault(r.SAMSUNG_CODE, []).append(dict(r))

    to_insert = []
    duplicates = []
    new_versions = []
    for r in rows:
        prior = existing_by_code.get(r["MODEL_CODE"])
        if not prior:
            to_insert.append(r)
            continue
        if any(str(p["QTY_ON_PROMOTION"]) == str(r["QTY"]) for p in prior):
            duplicates.append(r)
            continue
        to_insert.append(r)
        new_versions.append((r, prior[-1]))

    return to_insert, duplicates, new_versions

def push_promotion_rows(reference_no, title, target_table, to_insert, new_versions, sent_date=None):
    from google.cloud import bigquery
    client = get_bq_client()
    table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{target_table}"

    table = client.get_table(table_id)
    all_columns = [f.name for f in table.schema]

    if not to_insert:
        return 0, None

    prior_by_code = {r["MODEL_CODE"]: prior for r, prior in new_versions}
    upload_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc)

    records = []
    log_records = []
    for r in to_insert:
        total_value = round(r["PER_UNIT"] * r["QTY"], 2)
        rec = {c: None for c in all_columns}
        rec["MODEL"]            = title
        rec["PROMOTION_NO"]     = reference_no
        rec["SAMSUNG_CODE"]     = r["MODEL_CODE"]
        rec["START_DATE"]       = r["START_DATE"].strftime("%Y-%m-%d %H:%M:%S")
        rec["WEEK"]             = promo_week_str(r["START_DATE"])
        rec["END_DATE"]         = r["END_DATE"].strftime("%Y-%m-%d %H:%M:%S")
        rec["WEEK_2"]           = promo_week_str(r["END_DATE"])
        rec["QTY_ON_PROMOTION"] = str(r["QTY"])
        rec["PER_UNIT"]         = str(r["PER_UNIT"])
        rec["TOTAL_VALUE"]      = str(total_value)
        rec["MODEL_2"]          = r["PRODUCT"]
        rec["TOTAL_CLAIMED_SO_FAR"]    = "0"
        rec["CLAIMED_VALUE"]           = "0"
        rec["REMAINING_BALANCE"]       = str(r["QTY"])
        rec["REMAINING_CLAIM_VALUE"]   = str(total_value)
        rec["SENT_DATE"]        = sent_date.strftime("%Y-%m-%d %H:%M:%S") if sent_date else None
        rec["UPLOAD_ID"]        = upload_id
        records.append(rec)

        prior = prior_by_code.get(r["MODEL_CODE"])
        log_records.append({
            "LOG_ID": str(uuid.uuid4()),
            "UPLOAD_ID": upload_id,
            "TIMESTAMP": now,
            "TARGET_TABLE": target_table,
            "PROMOTION_NO": reference_no,
            "PROMOTION_TITLE": title,
            "SAMSUNG_CODE": r["MODEL_CODE"],
            "ACTION": "UPDATE" if prior else "INSERT",
            "OLD_QTY": str(prior["QTY_ON_PROMOTION"]) if prior else None,
            "NEW_QTY": str(r["QTY"]),
            "OLD_TOTAL_VALUE": str(prior["TOTAL_VALUE"]) if prior else None,
            "NEW_TOTAL_VALUE": str(total_value),
            "REVERTED": False,
            "REVERTED_AT": None,
        })

    df = pd.DataFrame(records, columns=all_columns)
    job_config = bigquery.LoadJobConfig(write_disposition=bigquery.WriteDisposition.WRITE_APPEND)
    job = client.load_table_from_dataframe(df, table_id, job_config=job_config)
    job.result()

    log_table_id = f"{BQ_PROJECT}.{BQ_DATASET}.promotion_logs"
    log_df = pd.DataFrame(log_records)
    log_job = client.load_table_from_dataframe(log_df, log_table_id, job_config=bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND
    ))
    log_job.result()

    return len(df), upload_id

# ══════════════════════════════════════════════════════
# PROMOTION SHEET SYNC — SKU coverage, extension & price-change detection
# ══════════════════════════════════════════════════════
PROMO_SHEET_TABS = {
    "promotion_data": 183516201,   # "PROMOTION DATA" tab
    "promotion_sm": 2031721001,    # "PROMOTION - SM" tab
}

# New highlight colors, following the same convention as the existing
# 18 Final Summary / Promotions color rules (a fixed color per condition,
# applied via background formatting / cell notes).
PROMO_HISTORY_COLORS = {
    "NEW_SKU": {"red": 1, "green": 1, "blue": 0.6},          # light yellow — newly synced row (existing convention)
    "EXTENSION": {"red": 0.80, "green": 0.89, "blue": 0.97}, # light blue — extends a prior promo for this SKU
    "PRICE_CHANGE": {"red": 0.96, "green": 0.80, "blue": 0.80}, # salmon — price changed vs prior promo for this SKU
}

EXTENSION_WINDOW_DAYS = 14  # a prior promo for the same SKU ending within this many days
                            # of this promo's start is treated as "extended"

@st.cache_data(ttl=600)
def get_models_sheet_df():
    """Full MODELS tab from the BACKWORKING sheet (MODEL / Sage Code / SAMSUNG CODE etc.)."""
    gc = get_gsheet_client()
    if gc is None:
        return pd.DataFrame()
    sh = gc.open_by_key(BACKWORKING_SHEET_ID)
    ws = sh.worksheet("MODELS")
    vals = ws.get_all_values()
    if not vals:
        return pd.DataFrame()
    return pd.DataFrame(vals[1:], columns=vals[0])

def expand_family_skus(samsung_code):
    """Given a SAMSUNG_CODE (e.g. SM-S938BZKKTPA), return every (sage_code, samsung_code, model)
    row in the MODELS sheet belonging to the same model family (matched by model-name prefix)."""
    code_to_model, _ = get_model_code_map()
    model = code_to_model.get(str(samsung_code).strip().upper())
    if not model:
        return []
    m = re.match(r'^([A-Z0-9]+)', str(model).strip())
    if not m:
        return []
    prefix = m.group(1)
    models_df = get_models_sheet_df()
    if models_df.empty or "MODEL" not in models_df.columns:
        return []
    fam = models_df[models_df["MODEL"].str.upper().str.startswith(prefix)]
    out = []
    for _, r in fam.iterrows():
        sc = str(r.get("SAMSUNG CODE", "")).strip()
        sg = str(r.get("Sage Code", "")).strip()
        if sc:
            out.append({"sage_code": sg, "samsung_code": sc, "model": r.get("MODEL", "")})
    return out

# Bulletins sometimes list a whole model family (e.g. "G-S25-ULTRA") instead of
# individual SKUs. Map each family code to the MODEL-name prefix(es) in the MODELS
# sheet so these can be expanded into real per-SKU rows before syncing to the sheet.
FAMILY_CODE_PREFIXES = {
    'G-A07': ['A075'], 'G-S26-ULTRA': ['S948'], 'G-A36-5G': ['A366'], 'G-A57-5G': ['A576'],
    'G-A37-5G': ['A376'], 'G-A56-5G': ['A566'], 'G-S26': ['S942'], 'G-S26+': ['S947'],
    'G-A17': ['A176'], 'G-S25-ULTRA': ['S938'], 'G-TAB-A11': ['TAB X135', 'TAB X133'],
    'G-WATCH8': ['L320', 'L330'], 'G-WATCH8-CLASSIC': ['L500'],
    'G-TAB-S10-LITE': ['TAB X400'],
}

def expand_family_code(family_code):
    """Given a family-level model code (e.g. G-S25-ULTRA), return every (sage_code,
    samsung_code, model) row in the MODELS sheet belonging to that family, via
    FAMILY_CODE_PREFIXES. Skips MODELS rows where Sage Code is blank or equal to the
    Samsung Code (duplicate/placeholder entries). Returns [] if the code isn't a known
    family code or has no resolvable SKU."""
    prefixes = FAMILY_CODE_PREFIXES.get(str(family_code).strip().upper())
    if not prefixes:
        return []
    models_df = get_models_sheet_df()
    if models_df.empty or "MODEL" not in models_df.columns:
        return []
    seen = set()
    out = []
    for _, r in models_df.iterrows():
        model = str(r.get("MODEL", "")).strip()
        sc = str(r.get("SAMSUNG CODE", "")).strip()
        sage = str(r.get("Sage Code", "")).strip()
        if not sc or not sage or sage.upper() == sc.upper() or sc.upper() in seen:
            continue
        if any(model.upper().startswith(p) for p in prefixes):
            seen.add(sc.upper())
            out.append({"sage_code": sage, "samsung_code": sc, "model": model})
    return out

def expand_family_rows_for_sync(to_insert):
    """Replace any to_insert row whose MODEL_CODE is a family-level code (per
    FAMILY_CODE_PREFIXES) with one row per real SKU in that family (per
    expand_family_code), with QTY split evenly across SKUs (remainder on first rows).
    Rows that aren't family codes, or have no resolvable SKU, pass through unchanged."""
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

def get_promo_price_history(samsung_codes, exclude_ref):
    """Most recent prior promo (across promotion_data + promotion_sm) per SAMSUNG_CODE,
    excluding `exclude_ref`. Returns {SAMSUNG_CODE: {PROMOTION_NO, START_DATE, END_DATE, PER_UNIT}}."""
    from google.cloud import bigquery
    if not samsung_codes:
        return {}
    client = get_bq_client()
    codes_params = [bigquery.ScalarQueryParameter(f"c{i}", "STRING", c) for i, c in enumerate(samsung_codes)]
    codes_list = ", ".join(f"@c{i}" for i in range(len(samsung_codes)))
    q = f"""
        SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, PER_UNIT FROM (
            SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, PER_UNIT,
                   ROW_NUMBER() OVER (PARTITION BY SAMSUNG_CODE ORDER BY END_DATE DESC) AS rn
            FROM (
                SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, PER_UNIT
                FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_data`
                WHERE SAMSUNG_CODE IN ({codes_list}) AND PROMOTION_NO != @ref
                UNION ALL
                SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, PER_UNIT
                FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_sm`
                WHERE SAMSUNG_CODE IN ({codes_list}) AND PROMOTION_NO != @ref
            )
        ) WHERE rn = 1
    """
    job = client.query(q, job_config=bigquery.QueryJobConfig(
        query_parameters=codes_params + [bigquery.ScalarQueryParameter("ref", "STRING", exclude_ref)]
    ))
    return {r.SAMSUNG_CODE: dict(r) for r in job}

def analyze_promo_changes(reference_no, target_table, to_insert):
    """For rows about to be inserted, detect (without writing anything):
    - extension of a prior promo for the same SKU (prior promo for this SKU ended within
      EXTENSION_WINDOW_DAYS of this promo's start date)
    - price change vs. the most recent prior promo for the same SKU
    - SKUs in the same model family (per MODELS sheet) that this promo doesn't cover at all

    Returns (per_row_flags, coverage_gaps).
    """
    codes = [r["MODEL_CODE"] for r in to_insert]
    history = get_promo_price_history(codes, reference_no)

    per_row = []
    for r in to_insert:
        code = r["MODEL_CODE"]
        flags = {"MODEL_CODE": code, "extension": False, "price_change": False}
        prior = history.get(code)
        if prior:
            prior_end = pd.to_datetime(prior["END_DATE"]).to_pydatetime().replace(tzinfo=None)
            prior_start = pd.to_datetime(prior["START_DATE"]).to_pydatetime().replace(tzinfo=None)
            gap_days = (r["START_DATE"] - prior_end).days
            if 0 <= gap_days <= EXTENSION_WINDOW_DAYS:
                flags["extension"] = True
                flags["prior_promo_no"] = prior["PROMOTION_NO"]
                flags["prior_start"] = prior_start.date()
                flags["prior_end"] = prior_end.date()
            old_price = to_num(prior["PER_UNIT"])
            if old_price is not None and abs(old_price - r["PER_UNIT"]) > 0.001:
                flags["price_change"] = True
                flags["old_price"] = old_price
                flags["new_price"] = r["PER_UNIT"]
        per_row.append(flags)

    covered_codes = {c.strip().upper() for c in codes}
    coverage_gaps = []
    seen_families = set()
    for code in codes:
        family = expand_family_skus(code)
        if not family:
            continue
        fam_key = tuple(sorted(f["samsung_code"].upper() for f in family))
        if fam_key in seen_families:
            continue
        seen_families.add(fam_key)
        missing = [f for f in family if f["samsung_code"].strip().upper() not in covered_codes]
        if missing:
            coverage_gaps.append({"sample_code": code, "missing_skus": missing})

    return per_row, coverage_gaps

def sync_promo_rows_to_sheet(reference_no, title, target_table, to_insert, per_row_flags, sent_date=None):
    """Append newly-inserted promo rows to the matching BACKWORKING sheet tab
    (PROMOTION DATA / PROMOTION - SM), highlighting new-SKU / extension / price-change
    rows per PROMO_HISTORY_COLORS, and adding cell notes for extensions and price changes.
    Column C (SAGE CODE, VLOOKUP'd from column D) is protected and is never written to.
    Any row whose MODEL_CODE is a family-level code (e.g. G-S25-ULTRA) is expanded into
    one row per real SKU in that family (per FAMILY_CODE_PREFIXES / MODELS sheet), with
    QTY equal-split across the SKUs, so the sheet only ever sees real Sage/Samsung codes.
    `sent_date` (the bulletin's own "Date Sent :" field) is added as a cell note on column A —
    no new column is added, since inserting one would shift every formula in the wide
    claim-tracking blocks to the right."""
    gid = PROMO_SHEET_TABS.get(target_table)
    if gid is None or not to_insert:
        return 0
    gc = get_gsheet_client()
    sh = gc.open_by_key(BACKWORKING_SHEET_ID)
    ws = sh.get_worksheet_by_id(gid)

    to_insert = expand_family_rows_for_sync(to_insert)
    flags_by_code = {f["MODEL_CODE"]: f for f in per_row_flags}
    start_row = len(ws.get_all_values()) + 1

    ab_rows, dj_rows, row_meta = [], [], []
    for i, r in enumerate(to_insert):
        row_no = start_row + i
        ab_rows.append([title, reference_no])
        dj_rows.append([
            r["MODEL_CODE"],
            r["START_DATE"].strftime("%m-%d-%y"),
            "",
            r["END_DATE"].strftime("%m-%d-%y"),
            "",
            r["QTY"],
            f"${r['PER_UNIT']:.2f}",
        ])
        row_meta.append((row_no, flags_by_code.get(r["MODEL_CODE"], {})))

    ws.update(f"A{start_row}:B{start_row + len(ab_rows) - 1}", ab_rows, value_input_option="RAW")
    # D:J only — K (TOTAL VALUE) is a formula (=I*J) and auto-fills, never written directly.
    ws.update(f"D{start_row}:J{start_row + len(dj_rows) - 1}", dj_rows, value_input_option="RAW")

    # Mirror the same rows to the other working sheets (same tab gids).
    # C is not protected there so VLOOKUP resolves the sage code from D automatically.
    import time as _time
    for mirror_name, mirror_id in (("Synced-sheet", SYNCED_SHEET_ID),
                                    ("Legacy-BACKWORKING", LEGACY_BACKWORKING_SHEET_ID)):
        if not mirror_id:
            continue
        try:
            sh2 = gc.open_by_key(mirror_id)
            ws2 = sh2.get_worksheet_by_id(gid)
            sr2 = len(ws2.get_all_values()) + 1
            end2 = sr2 + len(ab_rows) - 1
            _time.sleep(2)
            ws2.update(f"A{sr2}:B{end2}", ab_rows, value_input_option="RAW")
            _time.sleep(2)
            ws2.update(f"D{sr2}:J{end2}", dj_rows, value_input_option="RAW")
        except Exception as _e:
            st.warning(f"{mirror_name} mirror failed: {_e}")

    for row_no, flags in row_meta:
        color = PROMO_HISTORY_COLORS["NEW_SKU"]
        if flags.get("price_change"):
            color = PROMO_HISTORY_COLORS["PRICE_CHANGE"]
        if flags.get("extension"):
            color = PROMO_HISTORY_COLORS["EXTENSION"]
        ws.format(f"A{row_no}:B{row_no}", {"backgroundColor": color})
        ws.format(f"D{row_no}:J{row_no}", {"backgroundColor": color})
        if flags.get("extension"):
            ws.update_note(f"B{row_no}",
                f"Extension of promo {flags['prior_promo_no']} "
                f"(active {flags['prior_start']} to {flags['prior_end']})")
        if flags.get("price_change"):
            ws.update_note(f"J{row_no}",
                f"Price changed vs promo {flags.get('prior_promo_no', 'prior')}: "
                f"was ${flags['old_price']:.2f}, now ${flags['new_price']:.2f}")
        if sent_date:
            ws.update_note(f"A{row_no}", f"Date Sent: {sent_date.strftime('%Y-%m-%d')}")

    return len(to_insert)

def get_promotion_logs(limit=200):
    from google.cloud import bigquery
    client = get_bq_client()
    q = f"""
        SELECT * FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_logs`
        ORDER BY TIMESTAMP DESC
        LIMIT {limit}
    """
    return client.query(q).to_dataframe()

# ══════════════════════════════════════════════════════
# FINAL SUMMARY (live stock + IMEI + promotion status per model)
# ══════════════════════════════════════════════════════
@st.cache_data(ttl=1800)
def get_model_code_map():
    client = get_bq_client()
    df = client.query(f"SELECT MODEL, TYPE, SAGECODE, SAMSUNGCODE FROM `{BQ_PROJECT}.{BQ_DATASET}.report_models`").to_dataframe()
    code_to_model = {}
    model_type = {}
    for _, r in df.iterrows():
        model = r["MODEL"]
        if pd.notna(r["TYPE"]):
            model_type[model] = r["TYPE"]
        for code in [r["SAGECODE"], r["SAMSUNGCODE"]]:
            if code and pd.notna(code):
                code_to_model[str(code).strip().upper()] = model
    return code_to_model, model_type

@st.cache_data(ttl=600)
def get_sam_inventory_overrides(week_no):
    """Manually-entered 'SHOW SAM INVENTORY' values (Samsung-reported on-hand) for a week.
    Returns {model: qty}. Models with no override fall back to our computed INHAND."""
    from google.cloud import bigquery
    client = get_bq_client()
    df = client.query(f"""
        SELECT model, sam_inventory FROM `{BQ_PROJECT}.{BQ_DATASET}.sam_inventory_overrides`
        WHERE week_no = @w
    """, job_config=bigquery.QueryJobConfig(
        query_parameters=[bigquery.ScalarQueryParameter("w", "STRING", week_no)]
    )).to_dataframe()
    return dict(zip(df["model"], df["sam_inventory"]))

def save_sam_inventory_override(week_no, model, qty, updated_by="user"):
    from google.cloud import bigquery
    client = get_bq_client()
    delete_q = f"""
        DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.sam_inventory_overrides`
        WHERE week_no = @w AND model = @m
    """
    client.query(delete_q, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("w", "STRING", week_no),
        bigquery.ScalarQueryParameter("m", "STRING", model),
    ])).result()
    insert_q = f"""
        INSERT INTO `{BQ_PROJECT}.{BQ_DATASET}.sam_inventory_overrides`
        (week_no, model, sam_inventory, updated_by, updated_at)
        VALUES (@w, @m, @q, @u, CURRENT_TIMESTAMP())
    """
    client.query(insert_q, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("w", "STRING", week_no),
        bigquery.ScalarQueryParameter("m", "STRING", model),
        bigquery.ScalarQueryParameter("q", "INT64", int(qty)),
        bigquery.ScalarQueryParameter("u", "STRING", updated_by),
    ])).result()

# --- Google Sheet sync: approved manual edits get stored as overrides and
# applied on top of the computed Final Summary. ---

@st.cache_data(ttl=600)
def get_final_summary_overrides(week_no):
    """Approved manual edits from the BACKWORKING sheet: {(model, tab_name, column_name): value}."""
    from google.cloud import bigquery
    client = get_bq_client()
    df = client.query(f"""
        SELECT model, tab_name, column_name, override_value
        FROM `{BQ_PROJECT}.{BQ_DATASET}.final_summary_overrides`
        WHERE week_no = @w
    """, job_config=bigquery.QueryJobConfig(
        query_parameters=[bigquery.ScalarQueryParameter("w", "STRING", week_no)]
    )).to_dataframe()
    return {(r["model"], r["tab_name"], r["column_name"]): r["override_value"] for _, r in df.iterrows()}

def save_final_summary_override(week_no, model, tab_name, column_name, value, updated_by="user"):
    from google.cloud import bigquery
    client = get_bq_client()
    client.query(f"""
        DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.final_summary_overrides`
        WHERE week_no = @w AND model = @m AND tab_name = @t AND column_name = @c
    """, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("w", "STRING", week_no),
        bigquery.ScalarQueryParameter("m", "STRING", model),
        bigquery.ScalarQueryParameter("t", "STRING", tab_name),
        bigquery.ScalarQueryParameter("c", "STRING", column_name),
    ])).result()
    client.query(f"""
        INSERT INTO `{BQ_PROJECT}.{BQ_DATASET}.final_summary_overrides`
        (week_no, model, tab_name, column_name, override_value, updated_by, updated_at)
        VALUES (@w, @m, @t, @c, @v, @u, CURRENT_TIMESTAMP())
    """, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("w", "STRING", week_no),
        bigquery.ScalarQueryParameter("m", "STRING", model),
        bigquery.ScalarQueryParameter("t", "STRING", tab_name),
        bigquery.ScalarQueryParameter("c", "STRING", column_name),
        bigquery.ScalarQueryParameter("v", "STRING", str(value)),
        bigquery.ScalarQueryParameter("u", "STRING", updated_by),
    ])).result()

def _norm_cell(v):
    """Normalize a cell value (from sheet or computed df) for comparison."""
    if v is None:
        return ""
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return ""
    try:
        f = float(s.replace(",", ""))
        if f == int(f):
            return str(int(f))
        return f"{f:.2f}"
    except ValueError:
        return s

# Columns we don't compare / can't meaningfully push back (identifiers, derived flags).
NON_EDITABLE_COLS = {"MODEL", "TYPE", "ADJUSTED CHECK", "PROMO DATE CHECK", "PROMO ENDING (DAYS)"}

def detect_sheet_edits(week_no, rm_df, sm_df):
    """Compare the BACKWORKING sheet's synced tabs against the computed Final Summary.
    Returns a list of dicts describing cells that were manually changed in the sheet."""
    gc = get_gsheet_client()
    if gc is None:
        return []
    sh = gc.open_by_key(BACKWORKING_SHEET_ID)
    diffs = []
    for tab_name, computed in (("FINAL SUMMARY (App Sync)", rm_df), ("SM - FINAL SUMMARY (App Sync)", sm_df)):
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            continue
        sheet_values = ws.get_all_values()
        if not sheet_values:
            continue
        header = sheet_values[0]
        sheet_df = pd.DataFrame(sheet_values[1:], columns=header)
        if "MODEL" not in sheet_df.columns or "MODEL" not in computed.columns:
            continue
        computed_by_model = computed.set_index("MODEL")
        for _, srow in sheet_df.iterrows():
            model = srow["MODEL"]
            if model not in computed_by_model.index:
                continue
            crow = computed_by_model.loc[model]
            for col in header:
                if col in NON_EDITABLE_COLS or col not in computed.columns:
                    continue
                sheet_val = _norm_cell(srow[col])
                computed_val = _norm_cell(crow[col])
                if sheet_val != computed_val:
                    diffs.append({
                        "tab": tab_name,
                        "model": model,
                        "column": col,
                        "sheet_value": sheet_val,
                        "computed_value": computed_val,
                    })
    return diffs

@st.cache_data(ttl=300)
def get_pending_sheet_edits(week_no):
    from google.cloud import bigquery
    client = get_bq_client()
    return client.query(f"""
        SELECT edit_id, tab_name, model, column_name, sheet_value, computed_value
        FROM `{BQ_PROJECT}.{BQ_DATASET}.sheet_pending_edits`
        WHERE week_no = @w AND status = 'pending'
        ORDER BY detected_at
    """, job_config=bigquery.QueryJobConfig(
        query_parameters=[bigquery.ScalarQueryParameter("w", "STRING", week_no)]
    )).to_dataframe()

def record_pending_edits(week_no, diffs):
    """Insert newly-detected diffs as pending, skipping any already-pending with the same values."""
    from google.cloud import bigquery
    client = get_bq_client()
    existing = get_pending_sheet_edits(week_no)
    get_pending_sheet_edits.clear()
    existing_keys = set(zip(existing["tab_name"], existing["model"], existing["column_name"], existing["sheet_value"])) \
        if not existing.empty else set()
    rows = []
    for d in diffs:
        key = (d["tab"], d["model"], d["column"], d["sheet_value"])
        if key in existing_keys:
            continue
        rows.append({
            "edit_id": str(uuid.uuid4()),
            "week_no": week_no,
            "tab_name": d["tab"],
            "model": d["model"],
            "column_name": d["column"],
            "sheet_value": d["sheet_value"],
            "computed_value": d["computed_value"],
            "status": "pending",
            "detected_at": datetime.now(timezone.utc).isoformat(),
            "decided_by": None,
            "decided_at": None,
        })
    if rows:
        # DML insert (not insert_rows_json/streaming) so rows can be UPDATEd immediately on approve/reject.
        for r in rows:
            client.query(f"""
                INSERT INTO `{BQ_PROJECT}.{BQ_DATASET}.sheet_pending_edits`
                (edit_id, week_no, tab_name, model, column_name, sheet_value, computed_value, status, detected_at, decided_by, decided_at)
                VALUES (@edit_id, @week_no, @tab_name, @model, @column_name, @sheet_value, @computed_value, @status, CURRENT_TIMESTAMP(), NULL, NULL)
            """, job_config=bigquery.QueryJobConfig(query_parameters=[
                bigquery.ScalarQueryParameter("edit_id", "STRING", r["edit_id"]),
                bigquery.ScalarQueryParameter("week_no", "STRING", r["week_no"]),
                bigquery.ScalarQueryParameter("tab_name", "STRING", r["tab_name"]),
                bigquery.ScalarQueryParameter("model", "STRING", r["model"]),
                bigquery.ScalarQueryParameter("column_name", "STRING", r["column_name"]),
                bigquery.ScalarQueryParameter("sheet_value", "STRING", r["sheet_value"]),
                bigquery.ScalarQueryParameter("computed_value", "STRING", r["computed_value"]),
                bigquery.ScalarQueryParameter("status", "STRING", r["status"]),
            ])).result()
    return len(rows)

def resolve_pending_edit(edit_id, approve, week_no, model=None, tab_name=None, column=None, value=None, decided_by="user"):
    from google.cloud import bigquery
    client = get_bq_client()
    status = "approved" if approve else "rejected"
    client.query(f"""
        UPDATE `{BQ_PROJECT}.{BQ_DATASET}.sheet_pending_edits`
        SET status = @s, decided_by = @u, decided_at = CURRENT_TIMESTAMP()
        WHERE edit_id = @id
    """, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("s", "STRING", status),
        bigquery.ScalarQueryParameter("u", "STRING", decided_by),
        bigquery.ScalarQueryParameter("id", "STRING", edit_id),
    ])).result()
    if approve and model and column:
        save_final_summary_override(week_no, model, tab_name, column, value, updated_by=decided_by)
    get_pending_sheet_edits.clear()
    get_final_summary_overrides.clear()

@st.cache_data(ttl=600)
def get_summary_weeks():
    client = get_bq_client()
    df = client.query(f"""
        SELECT WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` WHERE WEEK_NO IS NOT NULL
        UNION DISTINCT
        SELECT WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE WEEK_NO IS NOT NULL
    """).to_dataframe()
    weeks = sorted(df["WEEK_NO"].dropna().unique().tolist(), key=week_to_num, reverse=True)
    return weeks

def prev_weeks(week_no, n):
    """Return [week_no, week_no-1, ..., week_no-(n-1)] as Y%4dW%02d strings (rolls back across years)."""
    m = re.match(r'Y(\d{4})W(\d+)', str(week_no))
    if not m:
        return [week_no]
    year, wk = int(m.group(1)), int(m.group(2))
    out = []
    for i in range(n):
        w = wk - i
        y = year
        while w < 1:
            y -= 1
            w += 52
        out.append(f"Y{y}W{w:02d}")
    return out

def weeks_strictly_between(start_week, end_week):
    """Weeks after start_week and before end_week (both exclusive), Y%4dW%02d, assumes 52 weeks/year."""
    sm = re.match(r'Y(\d{4})W(\d+)', str(start_week))
    em = re.match(r'Y(\d{4})W(\d+)', str(end_week))
    if not sm or not em:
        return []
    y, w = int(sm.group(1)), int(sm.group(2))
    end_y, end_w = int(em.group(1)), int(em.group(2))
    out = []
    while True:
        w += 1
        if w > 52:
            y += 1
            w = 1
        if (y, w) >= (end_y, end_w):
            break
        out.append(f"Y{y}W{w:02d}")
        if len(out) > 200:
            break
    return out

@st.cache_data(ttl=600)
def build_final_summary(week_no):
    from google.cloud import bigquery
    client = get_bq_client()
    code_to_model, model_type = get_model_code_map()
    sam_inv_overrides = get_sam_inventory_overrides(week_no)

    def map_model(code):
        if code is None or pd.isna(code):
            return None
        return code_to_model.get(str(code).strip().upper())

    week_param = bigquery.QueryJobConfig(query_parameters=[bigquery.ScalarQueryParameter("w", "STRING", week_no)])
    last4 = prev_weeks(week_no, 4)
    last4_param = bigquery.QueryJobConfig(query_parameters=[bigquery.ArrayQueryParameter("ws", "STRING", last4)])
    last8 = prev_weeks(week_no, 8)
    last8_param = bigquery.QueryJobConfig(query_parameters=[bigquery.ArrayQueryParameter("ws", "STRING", last8)])

    purch = client.query(f"""
        SELECT MATL, EIN FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` WHERE WEEK_NO = @w
    """, job_config=week_param).to_dataframe()

    # Aggregate in BigQuery instead of pulling all 2M+ purchase rows to pandas.
    purch_all = client.query(f"""
        SELECT MATL, COUNT(DISTINCT EIN) AS IMEI_COUNT
        FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases`
        GROUP BY MATL
    """).to_dataframe()

    # `sales` and `actual_sales` overlap heavily on IMEI (ACTUAL_SALE_IMEIS is
    # mostly a subset of SALE_DATABASE) — dedupe by IMEI so a sale isn't
    # counted twice when it appears in both tables.
    sales = client.query(f"""
        SELECT SKU, CUST_NO, WEEK_NO FROM (
            SELECT IMEI, SKU, CUST_NO, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE WEEK_NO = @w
            UNION ALL
            SELECT IMEI, SKU, CUST_NO, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales` WHERE WEEK_NO = @w
        )
        QUALIFY ROW_NUMBER() OVER (PARTITION BY IMEI ORDER BY IMEI) = 1
    """, job_config=week_param).to_dataframe()

    # Raw (non-deduped) sale rows for this week, used to flag models where the
    # same IMEI/sale appears more than once across sales/actual_sales — i.e.
    # "SALE COMMENTS & IMEI NOT MATCHED" needing human review.
    sales_raw = client.query(f"""
        SELECT SKU, COUNT(*) AS CNT FROM (
            SELECT IMEI, SKU FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE WEEK_NO = @w
            UNION ALL
            SELECT IMEI, SKU FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales` WHERE WEEK_NO = @w
        )
        GROUP BY SKU
    """, job_config=week_param).to_dataframe()
    sales_raw["MODEL"] = sales_raw["SKU"].map(map_model)
    raw_sale_count_by_model = sales_raw.groupby("MODEL")["CNT"].sum()

    sales_4w = client.query(f"""
        SELECT SKU, WEEK_NO FROM (
            SELECT IMEI, SKU, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE WEEK_NO IN UNNEST(@ws)
            UNION ALL
            SELECT IMEI, SKU, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales` WHERE WEEK_NO IN UNNEST(@ws)
        )
        QUALIFY ROW_NUMBER() OVER (PARTITION BY IMEI, WEEK_NO ORDER BY IMEI) = 1
    """, job_config=last4_param).to_dataframe()

    # Roll the opening balance forward from the latest available baseline week,
    # since opening_balance is only populated for Y2026W16 and looking it up
    # directly for later weeks returns 0 (causing artificial negative BALANCE).
    ob_weeks = client.query(f"""
        SELECT DISTINCT week_ref FROM `{BQ_PROJECT}.{BQ_DATASET}.opening_balance`
    """).to_dataframe()
    candidate_weeks = [w for w in ob_weeks["week_ref"].tolist() if week_to_num(w) <= week_to_num(week_no)]
    baseline_week = max(candidate_weeks, key=week_to_num) if candidate_weeks else None

    if baseline_week is not None:
        baseline_param = bigquery.QueryJobConfig(query_parameters=[bigquery.ScalarQueryParameter("w", "STRING", baseline_week)])
        ob = client.query(f"""
            SELECT sage_code, samsung_code, opening_balance
            FROM `{BQ_PROJECT}.{BQ_DATASET}.opening_balance` WHERE week_ref = @w
        """, job_config=baseline_param).to_dataframe()
        # NOTE: the "opening_balance" column in this table is, for some rows, a
        # nested dict (an artifact of how the table was loaded) whose own
        # "opening_balance" field holds the real integer — unwrap it here.
        ob["opening_balance"] = ob["opening_balance"].apply(
            lambda v: v.get("opening_balance") if isinstance(v, dict) else v
        )
    else:
        ob = pd.DataFrame(columns=["sage_code", "samsung_code", "opening_balance"])

    # opening_balance represents stock at the START of the baseline week, so the
    # baseline week's own purchases/sales must be rolled forward too — include
    # baseline_week itself in the delta range (not just the weeks after it).
    delta_weeks = ([baseline_week] + weeks_strictly_between(baseline_week, week_no)) if (baseline_week is not None and baseline_week != week_no) else []
    if delta_weeks:
        delta_param = bigquery.QueryJobConfig(query_parameters=[bigquery.ArrayQueryParameter("ws", "STRING", delta_weeks)])
        delta_purch = client.query(f"""
            SELECT MATL, COUNT(DISTINCT EIN) AS QTY
            FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` WHERE WEEK_NO IN UNNEST(@ws)
            GROUP BY MATL
        """, job_config=delta_param).to_dataframe()
        delta_sales = client.query(f"""
            SELECT SKU, COUNT(*) AS QTY FROM (
                SELECT IMEI, SKU FROM (
                    SELECT IMEI, SKU FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE WEEK_NO IN UNNEST(@ws)
                    UNION ALL
                    SELECT IMEI, SKU FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales` WHERE WEEK_NO IN UNNEST(@ws)
                )
                QUALIFY ROW_NUMBER() OVER (PARTITION BY IMEI ORDER BY IMEI) = 1
            ) GROUP BY SKU
        """, job_config=delta_param).to_dataframe()
        delta_purch["MODEL"] = delta_purch["MATL"].map(map_model)
        delta_sales["MODEL"] = delta_sales["SKU"].map(map_model)
        delta_purch_by_model = delta_purch.groupby("MODEL")["QTY"].sum()
        delta_sales_by_model = delta_sales.groupby("MODEL")["QTY"].sum()
    else:
        delta_purch_by_model = pd.Series(dtype="int64")
        delta_sales_by_model = pd.Series(dtype="int64")

    promo_all = client.query(f"""
        SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, QTY_ON_PROMOTION, PER_UNIT,
               TOTAL_CLAIMED_SO_FAR, CLAIMED_VALUE, REMAINING_BALANCE, REMAINING_CLAIM_VALUE,
               TOTAL_VALUE, TYPE, 'RM' AS SOURCE
        FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_data`
        WHERE SAMSUNG_CODE IS NOT NULL
        UNION ALL
        SELECT SAMSUNG_CODE, PROMOTION_NO, START_DATE, END_DATE, QTY_ON_PROMOTION, PER_UNIT,
               TOTAL_CLAIMED_SO_FAR, CLAIMED_VALUE, REMAINING_BALANCE, REMAINING_CLAIM_VALUE,
               TOTAL_VALUE, CAST(NULL AS STRING) AS TYPE, 'SM' AS SOURCE
        FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_sm`
        WHERE SAMSUNG_CODE IS NOT NULL
    """).to_dataframe()

    combo = client.query(f"""
        SELECT MASTER_MODEL, CHILD_MODEL, QTY_SALE, SALE_TO_SHOW, MARKET
        FROM `{BQ_PROJECT}.{BQ_DATASET}.combo_model_define`
    """).to_dataframe()

    purch["MODEL"]     = purch["MATL"].map(map_model)
    purch_all["MODEL"] = purch_all["MATL"].map(map_model)
    imei_count_by_model = purch_all.groupby("MODEL")["IMEI_COUNT"].sum()
    sales["MODEL"]     = sales["SKU"].map(map_model)
    sales_4w["MODEL"]  = sales_4w["SKU"].map(map_model)
    ob["MODEL"]        = ob["samsung_code"].map(map_model)
    ob.loc[ob["MODEL"].isna(), "MODEL"] = ob.loc[ob["MODEL"].isna(), "sage_code"].map(map_model)

    # split sell-thru (TYPE != 'P') vs price protection (TYPE == 'P')
    promo_all["MODEL"]     = promo_all["SAMSUNG_CODE"].map(map_model)
    promo_all["END_DATE_DT"]   = pd.to_datetime(promo_all["END_DATE"], errors="coerce")
    promo_all["START_DATE_DT"] = pd.to_datetime(promo_all["START_DATE"], errors="coerce")
    promo = promo_all[promo_all["TYPE"] != "P"]
    pp    = promo_all[promo_all["TYPE"] == "P"]

    sales_4w_by_model = sales_4w.groupby("MODEL")
    eight_week_set = set(prev_weeks(week_no, 8))

    promo_8w = client.query(f"""
        SELECT SAMSUNG_CODE, PROMOTION_NO, END_DATE, WEEK, WEEK_2
        FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_data`
        WHERE WEEK_2 IN UNNEST(@ws) OR WEEK IN UNNEST(@ws)
        UNION ALL
        SELECT SAMSUNG_CODE, PROMOTION_NO, END_DATE, WEEK, WEEK_2
        FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_sm`
        WHERE WEEK_2 IN UNNEST(@ws) OR WEEK IN UNNEST(@ws)
    """, job_config=last8_param).to_dataframe()
    promo_8w["MODEL"] = promo_8w["SAMSUNG_CODE"].map(map_model)

    today = pd.Timestamp.now()
    rows = []
    for model in sorted(set(code_to_model.values())):
        purchase_qty = len(purch[purch["MODEL"] == model])
        imei_count   = int(imei_count_by_model.get(model, 0))
        actual_sale  = len(sales[sales["MODEL"] == model])
        baseline_ob  = pd.to_numeric(ob.loc[ob["MODEL"] == model, "opening_balance"], errors="coerce").sum()
        inhand       = baseline_ob + delta_purch_by_model.get(model, 0) - delta_sales_by_model.get(model, 0)
        avail        = inhand + purchase_qty
        balance      = avail - actual_sale

        pm  = promo[promo["MODEL"] == model]
        ppm = pp[pp["MODEL"] == model]

        # ── Sell-thru promotion ──
        # picks the "best" (latest-ending active, else latest overall) promo from a
        # promo subset, then sums qty/remaining/claimed across all SAMSUNG_CODE
        # color/SKU variants that share that same PROMOTION_NO + END_DATE.
        def pick_promo(subset):
            if subset.empty:
                return None
            active_sub = subset[subset["END_DATE_DT"] >= today]
            # tie-break: latest END_DATE, then latest START_DATE (most recently
            # opened promo), then lowest PROMOTION_NO (earliest-created)
            pick_from = active_sub if not active_sub.empty else subset
            best_row = pick_from.sort_values(
                ["END_DATE_DT", "START_DATE_DT", "PROMOTION_NO"],
                ascending=[False, False, True],
            ).iloc[0]
            same = subset[(subset["PROMOTION_NO"] == best_row["PROMOTION_NO"])
                           & (subset["END_DATE_DT"] == best_row["END_DATE_DT"])]
            return {
                "no": best_row["PROMOTION_NO"],
                "end": best_row["END_DATE_DT"].strftime("%Y-%m-%d") if pd.notna(best_row["END_DATE_DT"]) else None,
                "start": best_row["START_DATE_DT"].strftime("%Y-%m-%d") if pd.notna(best_row["START_DATE_DT"]) else None,
                "value_pu": best_row["PER_UNIT"],
                "qty": sum(v for v in (to_num(x) for x in same["QTY_ON_PROMOTION"]) if v is not None),
                "remaining": sum(v for v in (to_num(x) for x in same["REMAINING_BALANCE"]) if v is not None),
                "claimed_qty": sum(v for v in (to_num(x) for x in same["TOTAL_CLAIMED_SO_FAR"]) if v is not None),
                "claimed_val": sum(v for v in (to_num(x) for x in same["CLAIMED_VALUE"]) if v is not None),
                "active": not active_sub.empty,
                "start_dt": best_row["START_DATE_DT"],
                "end_dt": best_row["END_DATE_DT"],
            }

        promo_no = promo_end = promo_start = promo_qty = promo_remaining = promo_value_pu = None
        promo_claimed_qty = promo_claimed_val = None
        promo_status = ""
        promo_ending_in_days = None
        upcoming_promo = "YES" if not pm.empty and (pm["START_DATE_DT"] > today).any() else ""
        promo_date_check = ""

        sm_promo_no = sm_promo_end = sm_promo_qty = sm_promo_remaining = None
        sm_promo_claimed_qty = sm_promo_claimed_val = None
        sm_promo_status = ""

        # ── RM vs SM promo (each picked independently within its own source) ──
        # FINAL SUMMARY (this sheet) is RM-sourced data, so the primary promo
        # fields below mirror the RM-only "best" promo to match that tab; the
        # SM-only "best" promo backs the separate SM PROMO ... columns, which
        # mirror the SM - FINAL SUMMARY tab.
        rm_best = pick_promo(pm[pm["SOURCE"] == "RM"])
        sm_best = pick_promo(pm[pm["SOURCE"] == "SM"])

        if rm_best is not None:
            promo_no          = rm_best["no"]
            promo_end         = rm_best["end"]
            promo_start       = rm_best["start"]
            promo_value_pu    = rm_best["value_pu"]
            promo_qty         = rm_best["qty"]
            promo_remaining   = rm_best["remaining"]
            promo_claimed_qty = rm_best["claimed_qty"]
            promo_claimed_val = rm_best["claimed_val"]
            promo_status      = "ACTIVE" if rm_best["active"] else "EXPIRED"
            if pd.notna(rm_best["end_dt"]):
                promo_ending_in_days = (rm_best["end_dt"] - today).days
            if pd.notna(rm_best["start_dt"]) and pd.notna(rm_best["end_dt"]) and rm_best["start_dt"] > rm_best["end_dt"]:
                promo_date_check = "YES"

        if sm_best is not None:
            sm_promo_no          = sm_best["no"]
            sm_promo_end         = sm_best["end"]
            sm_promo_qty         = sm_best["qty"]
            sm_promo_remaining   = sm_best["remaining"]
            sm_promo_claimed_qty = sm_best["claimed_qty"]
            sm_promo_claimed_val = sm_best["claimed_val"]
            sm_promo_status      = "ACTIVE" if sm_best["active"] else "EXPIRED"

        rm_per_unit = to_num(rm_best["value_pu"]) if rm_best else None
        sm_per_unit = to_num(sm_best["value_pu"]) if sm_best else None
        take_sm_promo = "YES" if (rm_per_unit is not None and sm_per_unit is not None
                                   and sm_per_unit > rm_per_unit) else ""

        # ── Price protection ──
        pp_start = pp_end = pp_claimable_qty = pp_total_value = pp_claimed_qty = pp_claimed_val = None
        if not ppm.empty:
            ppbest = ppm.sort_values("END_DATE_DT", ascending=False).iloc[0]
            pp_start = ppbest["START_DATE_DT"].strftime("%Y-%m-%d") if pd.notna(ppbest["START_DATE_DT"]) else None
            pp_end   = ppbest["END_DATE_DT"].strftime("%Y-%m-%d") if pd.notna(ppbest["END_DATE_DT"]) else None
            pp_claimable_qty = ppbest["QTY_ON_PROMOTION"]
            pp_total_value   = ppbest["TOTAL_VALUE"]
            pp_claimed_qty   = ppbest["TOTAL_CLAIMED_SO_FAR"]
            pp_claimed_val   = ppbest["CLAIMED_VALUE"]

        # ── Sale history ──
        m4 = sales_4w_by_model.get_group(model) if model in sales_4w_by_model.groups else pd.DataFrame(columns=["WEEK_NO"])
        last_4w_sale = len(m4)
        last_week_sale = (m4["WEEK_NO"] == prev_weeks(week_no, 2)[-1]).sum() if len(prev_weeks(week_no, 2)) > 1 else 0
        avg_weekly_sale = last_4w_sale / 4.0
        wos = round(balance / avg_weekly_sale, 1) if avg_weekly_sale > 0 else (None if balance == 0 else float("inf"))

        # ── No current promo but had one in last 8 weeks ──
        had_recent_promo = "NO"
        if promo_status != "ACTIVE":
            recent = promo_8w[(promo_8w["MODEL"] == model)]
            if not recent.empty and (set(recent["WEEK"]).union(set(recent["WEEK_2"])) & eight_week_set):
                had_recent_promo = "YES"

        # ── Combo model ──
        as_master = combo[combo["MASTER_MODEL"].str.contains(re.escape(model[:6]), case=False, na=False)] if not combo.empty else pd.DataFrame()
        as_child  = combo[combo["CHILD_MODEL"].str.contains(re.escape(model[:6]), case=False, na=False)] if not combo.empty else pd.DataFrame()
        master_combo_model = child_combo_model = combo_qty = combo_check = None
        if not as_master.empty:
            r0 = as_master.iloc[0]
            master_combo_model = r0["MASTER_MODEL"]
            child_combo_model  = r0["CHILD_MODEL"]
            combo_qty          = r0["SALE_TO_SHOW"]
            combo_check        = "OK" if pd.notna(combo_qty) and actual_sale >= float(combo_qty or 0) else "CHECK"
        elif not as_child.empty:
            r0 = as_child.iloc[0]
            master_combo_model = r0["MASTER_MODEL"]
            child_combo_model  = r0["CHILD_MODEL"]
            combo_qty          = r0["SALE_TO_SHOW"]

        # ── Promotion source ──
        if not pm.empty and promo_status == "ACTIVE":
            promo_source = "Bundle (Combo)" if not as_child.empty else "Sell-Thru"
        elif not ppm.empty:
            promo_source = "Price Protection"
        else:
            promo_source = "—"

        # ── SAM inventory / adjusted purchase check ──
        # SHOW SAM INVENTORY = Samsung-reported on-hand for this model. Defaults to
        # our computed INHAND (fully automatic). If someone enters a manual override
        # (because Samsung's number differs from ours), ADJUSTED and the CHK flag
        # reflect that discrepancy for human review.
        show_sam_inventory = sam_inv_overrides.get(model, inhand)
        adjusted = int(avail - show_sam_inventory)
        if show_sam_inventory == inhand:
            chk_flag = ""
        elif adjusted >= 0:
            chk_flag = "CHK MODEL SALES"
        else:
            chk_flag = "CHK SKU SALE"
        max_real_composite = f"{int(avail)} / {int(balance)} / {int(balance)} / {int(actual_sale)} / {int(last_4w_sale)}"

        # ── Sale/IMEI mismatch (raw row count vs deduped ACTUAL SALE count) ──
        sale_imei_check = "CHECK" if raw_sale_count_by_model.get(model, 0) != actual_sale else ""

        # ── On-hand vs remaining promo qty ──
        onhand_vs_promo_remaining = None
        promo_remaining_num = to_num(promo_remaining)
        if promo_remaining_num is not None:
            diff = int(promo_remaining_num) - int(show_sam_inventory)
            if diff > 0:
                onhand_vs_promo_remaining = diff

        rows.append({
            "MODEL": model,
            "TYPE": model_type.get(model, ""),
            "INHAND": int(inhand),
            "PURCHASE": purchase_qty,
            "AVAIL. FOR SALE": int(avail),
            "ACTUAL SALE": actual_sale,
            "BALANCE": int(balance),
            "MAX/REAL/MOST SALE/LAST WEEK SALE/BB": max_real_composite,
            "SHOW SAM INVENTORY": int(show_sam_inventory),
            "ADJUSTED": adjusted,
            "ADJUSTED CHECK": chk_flag,
            "SALE/IMEI CHECK": sale_imei_check,
            "IMEI COUNT": int(imei_count),
            "LAST WEEK SALE": int(last_week_sale),
            "LAST 3W+THIS SALE": int(last_4w_sale),
            "WOS": wos,
            "PROMOTION SOURCE": promo_source,
            "PROMO NUMBER": promo_no,
            "PROMO STATUS": promo_status,
            "UPCOMING PROMO": upcoming_promo,
            "PROMO START DATE": promo_start,
            "PROMO END DATE": promo_end,
            "PROMO DATE CHECK": promo_date_check,
            "PROMO ENDING (DAYS)": promo_ending_in_days,
            "NO PROMO BUT HAD IN LAST 8W": had_recent_promo,
            "QTY ON PROMOTION": promo_qty,
            "REMAINING BAL QTY": promo_remaining,
            "ONHAND < PROMO REMAINING": onhand_vs_promo_remaining,
            "VALUE P/U": promo_value_pu,
            "RM PROMO PER UNIT": rm_per_unit,
            "SM PROMO PER UNIT": sm_per_unit,
            "TAKE SM PROMO": take_sm_promo,
            "CLAIMED QTY": promo_claimed_qty,
            "CLAIMED VALUE": promo_claimed_val,
            "SM PROMO NUMBER": sm_promo_no,
            "SM PROMO STATUS": sm_promo_status,
            "SM PROMO END DATE": sm_promo_end,
            "SM QTY ON PROMOTION": sm_promo_qty,
            "SM REMAINING BAL QTY": sm_promo_remaining,
            "SM CLAIMED QTY": sm_promo_claimed_qty,
            "SM CLAIMED VALUE": sm_promo_claimed_val,
            "MASTER COMBO MODEL": master_combo_model,
            "CHILD COMBO MODEL": child_combo_model,
            "COMBO QTY": combo_qty,
            "COMBO CHECK": combo_check,
            "PP START DATE": pp_start,
            "PP END DATE": pp_end,
            "PP CLAIMABLE QTY": pp_claimable_qty,
            "PP TOTAL VALUE": pp_total_value,
            "PP CLAIMED QTY": pp_claimed_qty,
            "PP CLAIMED VALUE": pp_claimed_val,
        })

    overrides = get_final_summary_overrides(week_no)
    if overrides:
        for row in rows:
            model = row["MODEL"]
            for (m, tab_name, col), val in overrides.items():
                if m != model:
                    continue
                target_col = SM_RENAME_MAP_INV.get(col, col) if tab_name == SM_SYNC_TAB else col
                if target_col not in row:
                    continue
                try:
                    row[target_col] = float(val) if "." in str(val) else int(val)
                except (ValueError, TypeError):
                    row[target_col] = val

    return pd.DataFrame(rows)

# Exact header fill colors copied from the "FINAL SUMMARY" sheet in
# SAMSUNG WEEKLY - BACKWORKING.xlsx (per the owner's request to match exactly)
FINAL_SUMMARY_COLORS = {
    "MODEL": "#F1C232",
    "TYPE": "#F1C232",
    "INHAND": "#A4C2F4",
    "PURCHASE": "#F4CCCC",
    "AVAIL. FOR SALE": "#D9EAD3",
    "ACTUAL SALE": "#EA9999",
    "BALANCE": "#FFE599",
    "MAX/REAL/MOST SALE/LAST WEEK SALE/BB": "#D9EAD3",
    "SHOW SAM INVENTORY": "#D9EAD3",
    "ADJUSTED": "#CC0000",
    "ADJUSTED CHECK": "#CC0000",
    "SALE/IMEI CHECK": "#F6B26B",
    "IMEI COUNT": "#CC0000",
    "LAST WEEK SALE": "#CC0000",
    "LAST 3W+THIS SALE": "#BF9000",
    "WOS": "#1155CC",
    "PROMOTION SOURCE": "#1155CC",
    "PROMO NUMBER": "#3C78D8",
    "PROMO STATUS": "#3C78D8",
    "UPCOMING PROMO": "#FF00FF",
    "PROMO START DATE": "#3C78D8",
    "PROMO END DATE": "#3C78D8",
    "PROMO DATE CHECK": "#CC0000",
    "PROMO ENDING (DAYS)": "#1155CC",
    "NO PROMO BUT HAD IN LAST 8W": "#1155CC",
    "QTY ON PROMOTION": "#3C78D8",
    "REMAINING BAL QTY": "#38761D",
    "ONHAND < PROMO REMAINING": "#783F04",
    "VALUE P/U": "#3C78D8",
    "RM PROMO PER UNIT": "#3C78D8",
    "SM PROMO PER UNIT": "#38761D",
    "TAKE SM PROMO": "#38761D",
    "CLAIMED QTY": "#CC0000",
    "CLAIMED VALUE": "#CC0000",
    "SM PROMO NUMBER": "#38761D",
    "SM PROMO STATUS": "#38761D",
    "SM PROMO END DATE": "#38761D",
    "SM QTY ON PROMOTION": "#38761D",
    "SM REMAINING BAL QTY": "#38761D",
    "SM CLAIMED QTY": "#38761D",
    "SM CLAIMED VALUE": "#38761D",
    "MASTER COMBO MODEL": "#1155CC",
    "CHILD COMBO MODEL": "#1155CC",
    "COMBO QTY": "#1155CC",
    "COMBO CHECK": "#1155CC",
    "PP START DATE": "#CC0000",
    "PP END DATE": "#CC0000",
    "PP CLAIMABLE QTY": "#CC0000",
    "PP TOTAL VALUE": "#CC0000",
    "PP CLAIMED QTY": "#CC0000",
    "PP CLAIMED VALUE": "#CC0000",
}

def _text_color_for(hex_color):
    """Pick black or white text based on background luminance for readability."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#111111" if luminance > 0.55 else "#FFFFFF"

def style_final_summary(df):
    def col_style(col):
        color = FINAL_SUMMARY_COLORS.get(col.name)
        if not color:
            return [""] * len(col)
        text = _text_color_for(color)
        return [f"background-color: {color}; color: {text}"] * len(col)
    styler = df.style.apply(col_style, axis=0)
    styler = styler.format({
        "WOS": "{:.1f}",
        "PROMO ENDING (DAYS)": "{:.0f}",
        "COMBO QTY": "{:.0f}",
    }, na_rep="—")
    if "PROMO STATUS" in df.columns:
        def status_style(val):
            if val == "ACTIVE":
                return "background-color: #93C47D; color: #111; font-weight: 600"
            if val == "EXPIRED":
                return "background-color: #E06666; color: #111"
            return ""
        styler = styler.map(status_style, subset=["PROMO STATUS"])
        if "SM PROMO STATUS" in df.columns:
            styler = styler.map(status_style, subset=["SM PROMO STATUS"])

    # ── Conditional-formatting rules (per the "CONDITIONAL FORMATING COLOR
    # CLASSIFICATION" legend) — each row can light up specific cells based on
    # the values computed in build_final_summary().
    cols = list(df.columns)

    def row_style(row):
        styles = [""] * len(cols)

        def set_style(colname, css):
            if colname in cols:
                styles[cols.index(colname)] = css

        # 1. UPCOMING PROMO CHECK — magenta
        if row.get("UPCOMING PROMO") == "YES":
            set_style("UPCOMING PROMO", "background-color: #FF00FF; color: #111; font-weight: 600")

        # 2. PROMO ENDING IN DAYS FROM REPORTING DATE — yellow (<=7 days)
        pe = to_num(row.get("PROMO ENDING (DAYS)"))
        if pe is not None and 0 <= pe <= 7:
            set_style("PROMO ENDING (DAYS)", "background-color: #FFFF00; color: #111; font-weight: 600")

        # 3. COMBO PROMO CHILD MODEL — mauve
        if row.get("CHILD COMBO MODEL"):
            set_style("CHILD COMBO MODEL", "background-color: #C27BA0; color: #fff")

        # 4. MASTER COMBO MODEL — dark maroon
        if row.get("MASTER COMBO MODEL"):
            set_style("MASTER COMBO MODEL", "background-color: #741B47; color: #fff")

        # 5. PROMO MODEL — green text on white
        if row.get("PROMO STATUS") == "ACTIVE":
            set_style("MODEL", "background-color: #fff; color: #38761D; font-weight: 700")

        # 6. PROMO DATE CHECK — red text on black
        if row.get("PROMO DATE CHECK") == "YES":
            set_style("PROMO END DATE", "background-color: #000000; color: #FF0000; font-weight: 700")

        # 7. PRICE PROTECTION - CLAIMABLE QTY — white text on black
        ppq = to_num(row.get("PP CLAIMABLE QTY"))
        if ppq is not None and ppq > 0:
            set_style("PP CLAIMABLE QTY", "background-color: #000000; color: #FFFFFF; font-weight: 600")

        # 8. SALE COMMENTS & IMEI NOT MATCHED — orange
        if row.get("SALE/IMEI CHECK") == "CHECK":
            set_style("SALE/IMEI CHECK", "background-color: #F6B26B; color: #111; font-weight: 600")

        rem = to_num(row.get("REMAINING BAL QTY"))
        # 9. PROMO REMAINING QTY CHECK — yellow text on black (active promo, balance used up)
        if row.get("PROMO STATUS") == "ACTIVE" and (rem is None or rem <= 0):
            set_style("REMAINING BAL QTY", "background-color: #000000; color: #FFFF00; font-weight: 700")
        # 10. PROMO EXPIRING AND QTY REMAINING — mauve/pink
        elif pe is not None and 0 <= pe <= 7 and rem is not None and rem > 0:
            set_style("REMAINING BAL QTY", "background-color: #D5A6BD; color: #111; font-weight: 600")
        # 17. QTY ON HAND < REMAINING QTY ON PROMOTION — brown italic
        elif row.get("ONHAND < PROMO REMAINING") is not None and pd.notna(row.get("ONHAND < PROMO REMAINING")):
            set_style("REMAINING BAL QTY", "background-color: #783F04; color: #fff; font-style: italic; font-weight: 600")
            set_style("ONHAND < PROMO REMAINING", "background-color: #783F04; color: #fff; font-style: italic; font-weight: 600")

        # 11. NO CURRENT PROMO ... BUT HAD PROMO IN LAST 8 WEEKS — dark red bg, yellow italic
        if row.get("NO PROMO BUT HAD IN LAST 8W") == "YES":
            set_style("NO PROMO BUT HAD IN LAST 8W", "background-color: #990000; color: #FFFF00; font-style: italic; font-weight: 600")

        # 12/13. CHK MODEL SALES / CHK SKU SALE
        ac = row.get("ADJUSTED CHECK")
        if ac == "CHK MODEL SALES":
            set_style("ADJUSTED CHECK", "background-color: #E06666; color: #111; font-style: italic; font-weight: 600")
        elif ac == "CHK SKU SALE":
            set_style("ADJUSTED CHECK", "background-color: #F6B26B; color: #111; font-style: italic; font-weight: 600")

        # 14. COMBO QTY MODEL PROMO CHECK — dark red
        if row.get("COMBO CHECK") == "CHECK":
            set_style("COMBO CHECK", "background-color: #990000; color: #fff; font-weight: 600")

        # 15. PROMO PER UNIT CHECK WITH BULLETIN INCENTIVE — red-orange italic
        vpu = to_num(row.get("VALUE P/U"))
        if row.get("PROMO STATUS") == "ACTIVE" and (vpu is None or vpu <= 0):
            set_style("VALUE P/U", "background-color: #E06666; color: #111; font-style: italic; font-weight: 600")

        # 16. INCENTIVE ON BULLETIN — MENTIONED FOR PROMO MODEL — blue italic
        if row.get("PROMOTION SOURCE") == "Sell-Thru" and row.get("PROMO STATUS") == "ACTIVE":
            set_style("PROMOTION SOURCE", "background-color: #4A86E8; color: #fff; font-style: italic; font-weight: 600")

        # 18. DO NOT TAKE RM PROMO, TAKE SM PROMO — PRICE IS GREATER — dark green
        if row.get("TAKE SM PROMO") == "YES":
            set_style("PROMOTION SOURCE", "background-color: #38761D; color: #fff; font-weight: 700")
            set_style("TAKE SM PROMO", "background-color: #38761D; color: #fff; font-weight: 700")

        return styles

    styler = styler.apply(row_style, axis=1)
    return styler

# ══════════════════════════════════════════════════════
# PROMOTIONS FRAME (live, colored — same style as Final Summary)
# ══════════════════════════════════════════════════════
@st.cache_data(ttl=600)
def build_promotions_frame():
    client = get_bq_client()
    code_to_model, model_type = get_model_code_map()

    cols = """
        MODEL, PROMOTION_NO, SAMSUNG_CODE, SAGE_CODE, START_DATE, END_DATE,
        QTY_ON_PROMOTION, PER_UNIT, TOTAL_VALUE, TOTAL_CLAIMED_SO_FAR,
        CLAIMED_VALUE, REMAINING_BALANCE, REMAINING_CLAIM_VALUE, SENT_DATE
    """
    query = f"""
        SELECT {cols}, 'Sell-Thru' AS SOURCE FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_data`
        WHERE PROMOTION_NO IS NOT NULL
        UNION ALL
        SELECT {cols}, 'Special Market' AS SOURCE FROM `{BQ_PROJECT}.{BQ_DATASET}.promotion_sm`
        WHERE PROMOTION_NO IS NOT NULL
    """
    df = client.query(query).to_dataframe()
    if df.empty:
        return df

    today = pd.Timestamp.now().normalize()
    end_dates = pd.to_datetime(df["END_DATE"], errors="coerce")

    def status_for(end_dt):
        if pd.isna(end_dt):
            return "UNKNOWN"
        return "ACTIVE" if end_dt.date() >= today.date() else "EXPIRED"

    df["MODEL NAME"] = df["SAMSUNG_CODE"].map(code_to_model).fillna(df["MODEL"])
    df["STATUS"] = end_dates.map(status_for)
    df["ENDING (DAYS)"] = (end_dates - today).dt.days

    for c in ["QTY_ON_PROMOTION", "PER_UNIT", "TOTAL_VALUE", "TOTAL_CLAIMED_SO_FAR",
              "CLAIMED_VALUE", "REMAINING_BALANCE", "REMAINING_CLAIM_VALUE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.rename(columns={
        "MODEL": "MODEL CODE",
        "MODEL NAME": "MODEL",
        "SAMSUNG_CODE": "SAMSUNG CODE",
        "SAGE_CODE": "SAGE CODE",
        "PROMOTION_NO": "PROMO NUMBER",
        "START_DATE": "START DATE",
        "END_DATE": "END DATE",
        "QTY_ON_PROMOTION": "QTY ON PROMOTION",
        "PER_UNIT": "VALUE P/U",
        "TOTAL_VALUE": "TOTAL VALUE",
        "TOTAL_CLAIMED_SO_FAR": "CLAIMED QTY",
        "CLAIMED_VALUE": "CLAIMED VALUE",
        "REMAINING_BALANCE": "REMAINING BAL QTY",
        "REMAINING_CLAIM_VALUE": "REMAINING CLAIM VALUE",
        "SENT_DATE": "SENT DATE",
    })

    ordered = ["SOURCE", "MODEL", "MODEL CODE", "SAMSUNG CODE", "SAGE CODE", "PROMO NUMBER",
               "SENT DATE", "START DATE", "END DATE", "STATUS", "ENDING (DAYS)", "QTY ON PROMOTION",
               "VALUE P/U", "TOTAL VALUE", "CLAIMED QTY", "CLAIMED VALUE",
               "REMAINING BAL QTY", "REMAINING CLAIM VALUE"]
    df = df[ordered]

    # Active/unknown promotions first, then most-recently-expired, so the
    # rows users care about aren't buried under thousands of old expired ones.
    status_order = {"ACTIVE": 0, "UNKNOWN": 1, "EXPIRED": 2}
    df = df.assign(_status_order=df["STATUS"].map(status_order))
    df = df.sort_values(["_status_order", "END DATE"], ascending=[True, False]).drop(columns="_status_order")
    return df.reset_index(drop=True)

# Header colors reused from FINAL_SUMMARY_COLORS palette for visual consistency
PROMOTIONS_COLORS = {
    "MODEL": "#F1C232",
    "MODEL CODE": "#F1C232",
    "SAMSUNG CODE": "#F1C232",
    "SAGE CODE": "#F1C232",
    "SOURCE": "#1155CC",
    "PROMO NUMBER": "#3C78D8",
    "SENT DATE": "#3C78D8",
    "START DATE": "#3C78D8",
    "END DATE": "#3C78D8",
    "STATUS": "#3C78D8",
    "ENDING (DAYS)": "#1155CC",
    "QTY ON PROMOTION": "#3C78D8",
    "VALUE P/U": "#3C78D8",
    "TOTAL VALUE": "#3C78D8",
    "CLAIMED QTY": "#CC0000",
    "CLAIMED VALUE": "#CC0000",
    "REMAINING BAL QTY": "#38761D",
    "REMAINING CLAIM VALUE": "#38761D",
}

def style_promotions(df):
    def col_style(col):
        color = PROMOTIONS_COLORS.get(col.name)
        if not color:
            return [""] * len(col)
        text = _text_color_for(color)
        return [f"background-color: {color}; color: {text}"] * len(col)
    styler = df.style.apply(col_style, axis=0)
    styler = styler.format({
        "ENDING (DAYS)": "{:.0f}",
        "QTY ON PROMOTION": "{:,.0f}",
        "VALUE P/U": "{:,.2f}",
        "TOTAL VALUE": "{:,.2f}",
        "CLAIMED VALUE": "{:,.2f}",
        "REMAINING BAL QTY": "{:,.0f}",
        "REMAINING CLAIM VALUE": "{:,.2f}",
    }, na_rep="—")
    def status_style(val):
        if val == "ACTIVE":
            return "background-color: #93C47D; color: #111; font-weight: 600"
        if val == "EXPIRED":
            return "background-color: #E06666; color: #111"
        return ""
    styler = styler.map(status_style, subset=["STATUS"])
    return styler

def build_promotions_list(df_promos):
    """Collapse the per-SKU promotions frame into one row per (PROMO NUMBER, SOURCE) —
    a scannable list of every promotion ever pushed, instead of thousands of SKU rows."""
    if df_promos.empty:
        return df_promos
    g = df_promos.groupby(["PROMO NUMBER", "SOURCE"], as_index=False).agg(
        **{
            "START DATE": ("START DATE", "min"),
            "END DATE": ("END DATE", "max"),
            "STATUS": ("STATUS", "first"),
            "ENDING (DAYS)": ("ENDING (DAYS)", "first"),
            "SKU COUNT": ("SAMSUNG CODE", "nunique"),
            "QTY ON PROMOTION": ("QTY ON PROMOTION", "sum"),
            "TOTAL VALUE": ("TOTAL VALUE", "sum"),
            "CLAIMED QTY": ("CLAIMED QTY", "sum"),
            "CLAIMED VALUE": ("CLAIMED VALUE", "sum"),
            "REMAINING BAL QTY": ("REMAINING BAL QTY", "sum"),
            "REMAINING CLAIM VALUE": ("REMAINING CLAIM VALUE", "sum"),
        }
    )
    status_order = {"ACTIVE": 0, "UNKNOWN": 1, "EXPIRED": 2}
    g = g.assign(_so=g["STATUS"].map(status_order))
    g = g.sort_values(["_so", "END DATE"], ascending=[True, False]).drop(columns="_so")
    cols = ["PROMO NUMBER", "SOURCE", "START DATE", "END DATE", "STATUS", "ENDING (DAYS)",
            "SKU COUNT", "QTY ON PROMOTION", "TOTAL VALUE", "CLAIMED QTY", "CLAIMED VALUE",
            "REMAINING BAL QTY", "REMAINING CLAIM VALUE"]
    return g[cols].reset_index(drop=True)

# ══════════════════════════════════════════════════════
# AI CHAT (agentic, full read access to BigQuery)
# ══════════════════════════════════════════════════════
CHAT_MODEL = "claude-opus-4-8"

CHAT_SYSTEM_PROMPT = f"""You are a data assistant for the Samsung Purchases Pipeline app.
You have tool access to a BigQuery dataset `{BQ_PROJECT}.{BQ_DATASET}` containing tables such as:
purchases, sales, actual_sales, current_week_sales, purchase_imeis, model_master, model_prices,
promotion_data, promotion_sm, promotion_logs, rac_imei_uploads, opening_balance, report_models,
erp_purchases, erp_sales, manual_sales.

PROJECT CONTEXT — Closing / IMEI Reconciliation (see the "Closing" tab for the live UI version):
- "Sold" IMEIs = union of `sales` and `actual_sales` (excluding rows where SKU = 'N/A'), deduped by IMEI.
- "Orphan sold IMEIs" = IMEIs that appear in sold but have no matching EIN in either `purchases` or
  `purchase_imeis` — these represent sales with no purchase record (missing purchase data). As of the
  last check this is 21,546 IMEIs (down from a naive 48,116 before `purchase_imeis` was joined in).
- "In hand" = received (`purchases`) but not sold.
- Reconciliation identity: total_received (purchases row count) = imeis_sold + imeis_in_hand.
- Duplicate IMEIs (same EIN appearing more than once in `purchases`) = 0, verified clean.
- KNOWN DATA CAVEAT: ~568,015 rows in `purchases` have WEEK_NO = NULL and PO_NO = NULL, all dated
  PurchaseDate = 2022-12-31 — this is one old historical/opening-balance bulk-load batch, NOT an
  ongoing issue. Current weeks (e.g. Y2026Wxx) are fully populated. If asked about "missing week"
  or "missing PO" data, explain this is from that one legacy batch, not recent uploads.
- "Stock transfer to Special Market" (regular market vs special market IMEI-level transfers) is
  NOT currently tracked anywhere in this dataset — if asked, say this data doesn't exist yet and
  would need to be supplied/added before it can be reported on.
- Promotions: `promotion_data` = Regular/Sell-Thru market, `promotion_sm` = Special Market. A
  promotion's ACTIVE/EXPIRED status is based on END_DATE vs today.

Rules:
- ALWAYS use the `list_tables` and `get_table_schema` tools to confirm exact table/column names before writing SQL — never guess column names.
- ALWAYS use the `run_sql` tool to answer any question about actual data (counts, lookups, IMEI/model search, promotions, etc). Never make up numbers or rows from memory.
- If the user asks you to change/update/delete/insert data, first use `run_sql` (SELECT) to find and show exactly which row(s) would be affected, then call `run_write_sql` with the precise INSERT/UPDATE/DELETE/MERGE statement. The user will be asked to confirm before it runs.
- `run_write_sql` only accepts INSERT/UPDATE/DELETE/MERGE — never DROP/TRUNCATE/ALTER/CREATE.
- If a query returns no rows, say so plainly — do not invent results.
- For reconciliation/orphan/closing questions, replicate the logic described above in PROJECT CONTEXT
  (e.g. exclude IMEIs found in `purchase_imeis` when computing orphans) so your numbers match the Closing tab.
- Keep answers concise and reference the actual table/column names you used.
"""

def _bq_tool_list_tables():
    client = get_bq_client()
    tables = client.list_tables(f"{BQ_PROJECT}.{BQ_DATASET}")
    return [t.table_id for t in tables]

def _bq_tool_get_schema(table_name):
    client = get_bq_client()
    table = client.get_table(f"{BQ_PROJECT}.{BQ_DATASET}.{table_name}")
    return [{"name": f.name, "type": f.field_type} for f in table.schema]

def _bq_tool_run_sql(query):
    q = query.strip()
    # Strip a single trailing semicolon
    if q.endswith(";"):
        q = q[:-1]
    first_word = re.sub(r"[^a-zA-Z]", "", q.split(None, 1)[0]) if q.split() else ""
    if first_word.upper() != "SELECT":
        return {"error": "Only SELECT queries are allowed."}
    client = get_bq_client()
    try:
        df = client.query(q).to_dataframe(max_results=200000)
    except Exception as e:
        return {"error": str(e)}
    try:
        st.session_state["_chat_pending_df"] = df
        st.session_state["_chat_pending_sql"] = q
    except Exception:
        pass
    result = {
        "row_count": len(df),
        "rows": df.head(200).to_dict(orient="records"),
    }
    if len(df) > 200:
        result["note"] = (
            f"Only the first 200 of {len(df)} rows are shown here. "
            "The full result set is available to the user as a CSV download in the chat UI."
        )
    return result

def _bq_ensure_write_log_table():
    from google.cloud import bigquery
    client = get_bq_client()
    table_id = f"{BQ_PROJECT}.{BQ_DATASET}.ai_chat_write_logs"
    try:
        client.get_table(table_id)
    except Exception:
        schema = [
            bigquery.SchemaField("LOG_ID", "STRING"),
            bigquery.SchemaField("TIMESTAMP", "TIMESTAMP"),
            bigquery.SchemaField("SQL_TEXT", "STRING"),
            bigquery.SchemaField("AFFECTED_ROWS", "INTEGER"),
            bigquery.SchemaField("STATUS", "STRING"),
            bigquery.SchemaField("ERROR", "STRING"),
        ]
        client.create_table(bigquery.Table(table_id, schema=schema))
    return table_id

WRITE_KEYWORDS = {"INSERT", "UPDATE", "DELETE", "MERGE"}
FORBIDDEN_KEYWORDS_RE = re.compile(r"\b(DROP|TRUNCATE|ALTER|CREATE|GRANT|REVOKE)\b", re.IGNORECASE)

def _bq_tool_run_write_sql(query):
    """Executes an INSERT/UPDATE/DELETE/MERGE statement and logs it to ai_chat_write_logs."""
    from google.cloud import bigquery
    q = query.strip()
    if q.endswith(";"):
        q = q[:-1]
    first_word = re.sub(r"[^a-zA-Z]", "", q.split(None, 1)[0]) if q.split() else ""
    if first_word.upper() not in WRITE_KEYWORDS:
        return {"error": "Only INSERT, UPDATE, DELETE, or MERGE statements are allowed via run_write_sql."}
    if FORBIDDEN_KEYWORDS_RE.search(q):
        return {"error": "Statement contains a forbidden keyword (DROP/TRUNCATE/ALTER/CREATE/GRANT/REVOKE)."}

    client = get_bq_client()
    log_table_id = _bq_ensure_write_log_table()
    log_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc)

    try:
        job = client.query(q)
        job.result()
        affected = job.num_dml_affected_rows or 0
        status, error = "SUCCESS", None
        result = {"affected_rows": affected, "status": "success"}
    except Exception as e:
        affected, status, error = 0, "ERROR", str(e)
        result = {"error": str(e)}

    log_df = pd.DataFrame([{
        "LOG_ID": log_id, "TIMESTAMP": now, "SQL_TEXT": q,
        "AFFECTED_ROWS": affected, "STATUS": status, "ERROR": error,
    }])
    client.load_table_from_dataframe(log_df, log_table_id, job_config=bigquery.LoadJobConfig(
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND
    )).result()

    return result

CHAT_TOOLS = [
    {
        "name": "list_tables",
        "description": f"List all table names in the {BQ_DATASET} BigQuery dataset.",
        "input_schema": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_table_schema",
        "description": "Get the column names and types for a given BigQuery table.",
        "input_schema": {
            "type": "object",
            "properties": {"table_name": {"type": "string", "description": "Table name (without project/dataset prefix)"}},
            "required": ["table_name"],
        },
    },
    {
        "name": "run_sql",
        "description": f"Run a read-only SELECT SQL query against `{BQ_PROJECT}.{BQ_DATASET}`. Always fully-qualify table names as `{BQ_PROJECT}.{BQ_DATASET}.<table>`. Returns up to 200 rows.",
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "A SELECT SQL query."}},
            "required": ["query"],
        },
    },
    {
        "name": "run_write_sql",
        "description": (
            f"Run an INSERT, UPDATE, DELETE, or MERGE statement against `{BQ_PROJECT}.{BQ_DATASET}` to make a "
            f"data change the user explicitly asked for. Always fully-qualify table names as "
            f"`{BQ_PROJECT}.{BQ_DATASET}.<table>`. Always confirm the exact rows affected with a SELECT (run_sql) "
            f"first if there's any ambiguity. This requires user confirmation before it runs and is logged to "
            f"`ai_chat_write_logs`."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "An INSERT/UPDATE/DELETE/MERGE SQL statement."}},
            "required": ["query"],
        },
    },
]

def _execute_chat_tool(name, tool_input):
    if name == "list_tables":
        return _bq_tool_list_tables()
    if name == "get_table_schema":
        return _bq_tool_get_schema(tool_input["table_name"])
    if name == "run_sql":
        return _bq_tool_run_sql(tool_input["query"])
    if name == "run_write_sql":
        return _bq_tool_run_write_sql(tool_input["query"])
    return {"error": f"Unknown tool: {name}"}

@st.cache_resource
def get_anthropic_client():
    from anthropic import Anthropic
    api_key = safe_secret("ANTHROPIC_API_KEY")
    if api_key:
        return Anthropic(api_key=api_key)
    return Anthropic()

def run_chat_agent(history, max_steps=8):
    """history: list of {"role": "user"/"assistant", "content": ...} in Anthropic format.
    Runs the agentic tool-use loop. Returns one of:
      ("final", text, messages)
      ("pending_write", {"write_block": block, "pending_results": [...]}, messages)
    `messages` always includes the latest assistant turn appended."""
    client = get_anthropic_client()
    messages = list(history)

    for _ in range(max_steps):
        response = client.messages.create(
            model=CHAT_MODEL,
            max_tokens=4096,
            thinking={"type": "adaptive"},
            system=CHAT_SYSTEM_PROMPT,
            tools=CHAT_TOOLS,
            messages=messages,
        )

        if response.stop_reason != "tool_use":
            text_parts = [b.text for b in response.content if b.type == "text"]
            return "final", "\n".join(text_parts), messages + [{"role": "assistant", "content": response.content}]

        messages.append({"role": "assistant", "content": response.content})

        write_block = next((b for b in response.content
                             if b.type == "tool_use" and b.name == "run_write_sql"), None)

        tool_results = []
        for block in response.content:
            if block.type == "tool_use" and block is not write_block:
                result = _execute_chat_tool(block.name, block.input)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(result, default=str),
                })

        if write_block is not None:
            return "pending_write", {"write_block": write_block, "pending_results": tool_results}, messages

        messages.append({"role": "user", "content": tool_results})

    return "final", "Reached the tool-use step limit without a final answer.", messages

def revert_promotion_upload(upload_id, target_table):
    """Delete the rows added by this upload and mark the related log entries reverted."""
    from google.cloud import bigquery
    client = get_bq_client()
    table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{target_table}"

    delete_q = f"DELETE FROM `{table_id}` WHERE UPLOAD_ID = @upload_id"
    client.query(delete_q, job_config=bigquery.QueryJobConfig(
        query_parameters=[bigquery.ScalarQueryParameter("upload_id", "STRING", upload_id)]
    )).result()

    log_table_id = f"{BQ_PROJECT}.{BQ_DATASET}.promotion_logs"
    update_q = f"""
        UPDATE `{log_table_id}`
        SET REVERTED = TRUE, REVERTED_AT = @now
        WHERE UPLOAD_ID = @upload_id
    """
    client.query(update_q, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("now", "TIMESTAMP", datetime.now(timezone.utc)),
        bigquery.ScalarQueryParameter("upload_id", "STRING", upload_id),
    ])).result()

# ══════════════════════════════════════════════════════
# CLOSING / IMEI RECONCILIATION (3.7 — in vs out, orphans, PO traceability)
# ══════════════════════════════════════════════════════
SOLD_CTE = f"""
sold AS (
  SELECT IMEI, ANY_VALUE(SKU) AS SKU, ANY_VALUE(INVOICE_NO) AS INVOICE_NO,
         ANY_VALUE(CUSTOMER) AS CUSTOMER, ANY_VALUE(DATE) AS SALE_DATE, ANY_VALUE(WEEK_NO) AS SALE_WEEK
  FROM (
    SELECT IMEI, SKU, INVOICE_NO, CUSTOMER, DATE, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.sales` WHERE SKU != 'N/A'
    UNION ALL
    SELECT IMEI, SKU, INVOICE_NO, CUSTOMER, DATE, WEEK_NO FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales` WHERE SKU != 'N/A'
  )
  GROUP BY IMEI
)
"""

@st.cache_data(ttl=3600)
def build_closing_overview():
    client = get_bq_client()
    q = f"""
    WITH {SOLD_CTE}
    SELECT
      (SELECT COUNT(*) FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases`) AS total_received,
      (SELECT COUNT(*) FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` WHERE PO_NO IS NULL OR PO_NO = '') AS blank_po,
      (SELECT COUNT(*) FROM sold s
         LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchases` p ON p.EIN = s.IMEI
         LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchase_imeis` pi ON pi.ein = s.IMEI
         WHERE p.EIN IS NULL AND pi.ein IS NULL) AS orphan_imeis,
      (SELECT COUNTIF(s.IMEI IS NOT NULL) FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` p LEFT JOIN sold s ON s.IMEI = p.EIN) AS imeis_sold,
      (SELECT COUNTIF(s.IMEI IS NULL) FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` p LEFT JOIN sold s ON s.IMEI = p.EIN) AS imeis_in_hand,
      (SELECT COUNT(*) FROM (SELECT EIN FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` GROUP BY EIN HAVING COUNT(*) > 1)) AS duplicate_imeis
    """
    return client.query(q).to_dataframe().iloc[0]

@st.cache_data(ttl=3600)
def get_orphan_imeis(sku_search=""):
    from google.cloud import bigquery
    client = get_bq_client()
    where = ""
    params = []
    if sku_search:
        where = "AND UPPER(s.SKU) LIKE @sku"
        params.append(bigquery.ScalarQueryParameter("sku", "STRING", f"%{sku_search.upper()}%"))
    q = f"""
    WITH {SOLD_CTE}
    SELECT s.IMEI, s.SKU, s.SALE_DATE, s.SALE_WEEK, s.INVOICE_NO, s.CUSTOMER
    FROM sold s
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchases` p ON p.EIN = s.IMEI
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchase_imeis` pi ON pi.ein = s.IMEI
    WHERE p.EIN IS NULL AND pi.ein IS NULL {where}
    ORDER BY s.SKU, s.IMEI
    LIMIT 5000
    """
    return client.query(q, job_config=bigquery.QueryJobConfig(query_parameters=params)).to_dataframe()

@st.cache_data(ttl=3600)
def get_orphan_summary_by_sku():
    client = get_bq_client()
    q = f"""
    WITH {SOLD_CTE}
    SELECT s.SKU, COUNT(*) AS orphan_count
    FROM sold s
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchases` p ON p.EIN = s.IMEI
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchase_imeis` pi ON pi.ein = s.IMEI
    WHERE p.EIN IS NULL AND pi.ein IS NULL
    GROUP BY s.SKU
    ORDER BY orphan_count DESC
    """
    return client.query(q).to_dataframe()

@st.cache_data(ttl=3600)
def build_po_traceability():
    client = get_bq_client()
    q = f"""
    WITH {SOLD_CTE}
    SELECT
      COALESCE(p.PO_NO, '(blank)') AS PO_NO,
      COUNT(*) AS imeis_received,
      COUNTIF(s.IMEI IS NOT NULL) AS imeis_sold,
      COUNT(*) - COUNTIF(s.IMEI IS NOT NULL) AS imeis_in_hand,
      ANY_VALUE(p.MATL) AS sample_matl,
      ANY_VALUE(p.WEEK_NO) AS sample_week
    FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` p
    LEFT JOIN sold s ON s.IMEI = p.EIN
    GROUP BY PO_NO
    ORDER BY imeis_received DESC
    """
    return client.query(q).to_dataframe()

@st.cache_data(ttl=3600)
def get_imei_trace(imei):
    """Full in/out trace for a single IMEI: purchase (PO/date) + sale (invoice/customer/date)."""
    from google.cloud import bigquery
    client = get_bq_client()
    q = f"""
    WITH {SOLD_CTE}
    SELECT
      COALESCE(p.EIN, pi.ein, s.IMEI) AS IMEI,
      COALESCE(p.PO_NO, pi.po_no) AS PO_NO,
      COALESCE(p.MATL, pi.matl) AS PURCHASED_SKU,
      COALESCE(p.WEEK_NO, pi.week_no) AS PURCHASE_WEEK,
      COALESCE(p.PurchaseDate, pi.purchase_date) AS PurchaseDate,
      s.SKU AS SOLD_SKU, s.INVOICE_NO, s.CUSTOMER, s.SALE_DATE, s.SALE_WEEK,
      CASE
        WHEN pi.ein IS NOT NULL THEN 'FOUND IN purchase_imeis (not purchases) - ' ||
          (CASE WHEN s.IMEI IS NULL THEN 'IN HAND' ELSE 'SOLD' END)
        WHEN p.EIN IS NULL THEN 'SOLD - NO PURCHASE RECORD (orphan)'
        WHEN s.IMEI IS NULL THEN 'IN HAND (received, not sold)'
        ELSE 'SOLD'
      END AS STATUS
    FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` p
    FULL OUTER JOIN sold s ON s.IMEI = p.EIN
    LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.purchase_imeis` pi
      ON pi.ein = COALESCE(p.EIN, s.IMEI)
    WHERE p.EIN = @imei OR s.IMEI = @imei OR pi.ein = @imei
    """
    return client.query(q, job_config=bigquery.QueryJobConfig(query_parameters=[
        bigquery.ScalarQueryParameter("imei", "STRING", str(imei).strip())
    ])).to_dataframe()

@st.cache_data(ttl=3600)
def get_po_imei_detail(po_no):
    from google.cloud import bigquery
    client = get_bq_client()
    po_filter = "p.PO_NO IS NULL OR p.PO_NO = ''" if po_no == "(blank)" else "p.PO_NO = @po_no"
    params = [] if po_no == "(blank)" else [bigquery.ScalarQueryParameter("po_no", "STRING", po_no)]
    q = f"""
    WITH {SOLD_CTE}
    SELECT p.PO_NO, p.MATL, p.EIN AS IMEI, p.WEEK_NO AS PURCHASE_WEEK, p.PurchaseDate,
           s.SKU AS SOLD_SKU, s.INVOICE_NO, s.CUSTOMER, s.SALE_DATE,
           CASE WHEN s.IMEI IS NULL THEN 'IN_HAND' ELSE 'SOLD' END AS STATUS
    FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases` p
    LEFT JOIN sold s ON s.IMEI = p.EIN
    WHERE {po_filter}
    ORDER BY STATUS, p.EIN
    LIMIT 5000
    """
    return client.query(q, job_config=bigquery.QueryJobConfig(query_parameters=params)).to_dataframe()

# ══════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════
st.markdown("# 📦 Samsung Purchases Pipeline")
st.divider()

tab_pipeline, tab_imei, tab_model, tab_promo, tab_logs, tab_summary, tab_promotions, tab_closing, tab_chat = st.tabs(
    ["Pipeline", "IMEI Search", "Model Search", "Promotion PDF", "Logs", "Final Summary", "Promotions", "Closing", "AI Chat"])

# ──────────────────────────────────────────────────────
# TAB 1: PIPELINE
# ──────────────────────────────────────────────────────
with tab_pipeline:
    left, right = st.columns([1, 1], gap="large")

    with left:
        st.markdown("### Inputs")

        invoice_file = st.file_uploader("Samsung Invoice (.xlsx)", type=["xlsx"],
                                         key="invoice", on_change=reset_pipeline)

        st.markdown("**Data Sheet** — ODO#→PO#, Week, Purchase Date")
        ds_mode = st.radio("Source", ["Upload Excel file", "Google Sheets link"],
                           horizontal=True, on_change=reset_pipeline)

        data_sheet_file = None
        sheets_url      = None

        if ds_mode == "Upload Excel file":
            data_sheet_file = st.file_uploader("Data Sheet (.xlsx)", type=["xlsx"],
                                                key="data_sheet", on_change=reset_pipeline)
        else:
            sheets_url = st.text_input("Paste Google Sheets URL (must be public/viewer)",
                                        key="sheets_url", on_change=reset_pipeline,
                                        placeholder="https://docs.google.com/spreadsheets/d/...")

        st.divider()

        ready = invoice_file and (data_sheet_file or sheets_url)
        process_btn = st.button("⚙️ Process Files", width="stretch", disabled=not ready)

        if process_btn and ready:
            reset_pipeline()
            with st.spinner("Processing..."):
                try:
                    # Load data sheet
                    if data_sheet_file:
                        po_lookup, week_no, purchase_date = load_data_sheet_from_file(data_sheet_file)
                    else:
                        po_lookup, week_no, purchase_date = load_data_sheet_from_url(sheets_url)

                    if not week_no:
                        st.warning("Week number not found in data sheet — check column name is WEEK_NO with format Y2026WXX")

                    invoice_file.seek(0)
                    df_raw = pd.read_excel(invoice_file, sheet_name="IMEI", dtype=str)
                    df_out = transform(df_raw, po_lookup, week_no, purchase_date)

                    st.session_state.df_out          = df_out
                    st.session_state.total_rows       = len(df_out)
                    st.session_state.matched_po       = int(df_out["PO_NO"].notna().sum())
                    st.session_state.week_no_used     = week_no
                    st.session_state.purchase_date_used = purchase_date
                    st.session_state.step1_done       = True
                except Exception as e:
                    st.session_state.step1_done   = False
                    st.session_state.step1_status = "error"
                    st.session_state.step1_msg    = str(e)

            # ── Auto-run Steps 2-5 after a successful Step 1 ──
            if st.session_state.step1_done:
                df = st.session_state.df_out
                with st.spinner("Pushing to staging..."):
                    try:
                        from google.cloud import bigquery
                        client = get_bq_client()
                        job_config = bigquery.LoadJobConfig(
                            write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
                            schema=[bigquery.SchemaField(c, "STRING") for c in BQ_SCHEMA],
                        )
                        job = client.load_table_from_dataframe(df, f"{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}", job_config=job_config)
                        job.result()
                        st.session_state.step2_done = True
                        st.session_state.step2_msg  = f"{len(df):,} rows uploaded."
                    except Exception as e:
                        st.session_state.step2_done = False
                        st.session_state.step2_msg  = str(e)

                if st.session_state.step2_done:
                    with st.spinner("Checking for duplicate IMEIs..."):
                        try:
                            client = get_bq_client()
                            query = f"""
                                SELECT s.*, CASE WHEN m.EIN IS NOT NULL THEN 'DUPLICATE' ELSE NULL END AS RESULT
                                FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` s
                                LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.{TBL_IMEIS}` m ON s.EIN = m.EIN
                            """
                            result_df = client.query(query).to_dataframe()
                            dups_df   = result_df[result_df["RESULT"].notna()]
                            st.session_state.step3_done    = True
                            st.session_state.duplicates_df = dups_df
                        except Exception as e:
                            st.session_state.step3_done = False
                            st.session_state.step3_msg  = str(e)

                # Only auto-continue to Master push + clear staging if there were
                # NO duplicate IMEIs — otherwise pause here for manual review.
                if st.session_state.step3_done and len(st.session_state.duplicates_df) == 0:
                    with st.spinner("Pushing to master table..."):
                        try:
                            client = get_bq_client()
                            query = f"""
                                INSERT INTO `{BQ_PROJECT}.{BQ_DATASET}.{TBL_MASTER}`
                                    (SO_NO, BILLING_NO, MATL, EIN, DUAL_IMEI_NO, PO_NO, WEEK_NO, PurchaseDate)
                                SELECT SO_NO, BILLING_NO, MATL, EIN, DUAL_IMEI_NO, PO_NO, WEEK_NO, PurchaseDate
                                FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` s
                                WHERE NOT EXISTS (
                                    SELECT 1 FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_IMEIS}` m WHERE m.EIN = s.EIN
                                )
                            """
                            job = client.query(query)
                            job.result()
                            st.session_state.step4_done = True
                            st.session_state.step4_msg  = f"{job.num_dml_affected_rows or 0:,} rows inserted."
                        except Exception as e:
                            st.session_state.step4_done = False
                            st.session_state.step4_msg  = str(e)

                    if st.session_state.step4_done:
                        with st.spinner("Clearing staging table..."):
                            try:
                                client = get_bq_client()
                                client.query(f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` WHERE TRUE").result()
                                st.session_state.step5_done = True
                                st.session_state.step5_msg  = "Staging cleared. Pipeline complete!"
                            except Exception as e:
                                st.session_state.step5_done = False
                                st.session_state.step5_msg  = str(e)
                elif st.session_state.step3_done and len(st.session_state.duplicates_df) > 0:
                    st.warning(f"⚠ Pipeline paused: {len(st.session_state.duplicates_df)} duplicate IMEI(s) found. "
                               f"Review below, then click 'Push to Master Table' to continue manually.")

        if st.session_state.df_out is not None:
            st.markdown("**Preview (5 rows)**")
            st.dataframe(st.session_state.df_out.head(5), width="stretch", hide_index=True)

            total   = st.session_state.total_rows
            matched = st.session_state.matched_po
            missing = total - matched
            week_used = st.session_state.week_no_used or "—"
            date_used = st.session_state.purchase_date_used or "—"

            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-box"><div class="metric-val">{total:,}</div><div class="metric-lbl">Total Rows</div></div>
                <div class="metric-box"><div class="metric-val" style="color:#00e676">{matched:,}</div><div class="metric-lbl">PO Matched</div></div>
                <div class="metric-box"><div class="metric-val" style="color:{'#ff1744' if missing>0 else '#00e676'}">{missing:,}</div><div class="metric-lbl">Missing PO</div></div>
            </div>
            <div style="font-family:IBM Plex Mono,monospace;font-size:0.8rem;color:#555;margin-top:0.5rem">
                Week: {week_used} &nbsp;|&nbsp; Purchase Date: {date_used}
            </div>
            """, unsafe_allow_html=True)

            buf = io.BytesIO()
            st.session_state.df_out.to_excel(buf, index=False)
            st.download_button("⬇ Download Preview Excel", buf.getvalue(),
                               file_name=f"purchases_{week_used}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               width="stretch")

    with right:
        st.markdown("### Pipeline")
        df = st.session_state.df_out

        # Step 1
        if st.session_state.step1_done is None:
            step_card(1, "Process Files", "idle", "Waiting for files.")
        elif st.session_state.step1_done:
            step_card(1, "Process Files", "success",
                      f"{st.session_state.total_rows:,} rows transformed.",
                      f"Week: {st.session_state.week_no_used} | PO matched: {st.session_state.matched_po:,}")
        else:
            step_card(1, "Process Files", "error", st.session_state.step1_msg or "Failed.")

        # Step 2
        step2_disabled = not st.session_state.step1_done
        if st.session_state.step2_done is None:
            step_card(2, "Upload to Staging", "idle", f"Push data to `{TBL_STAGING}`.")
        elif st.session_state.step2_done:
            step_card(2, "Upload to Staging", "success", st.session_state.step2_msg, f"{BQ_DATASET}.{TBL_STAGING}")
        else:
            step_card(2, "Upload to Staging", "error", st.session_state.step2_msg)

        if not step2_disabled and st.session_state.step2_done is None:
            if st.button("▶ Push to Staging", width="stretch", key="btn2"):
                with st.spinner("Uploading..."):
                    try:
                        from google.cloud import bigquery
                        client = get_bq_client()
                        job_config = bigquery.LoadJobConfig(
                            write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
                            schema=[bigquery.SchemaField(c, "STRING") for c in BQ_SCHEMA],
                        )
                        job = client.load_table_from_dataframe(df, f"{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}", job_config=job_config)
                        job.result()
                        st.session_state.step2_done = True
                        st.session_state.step2_msg  = f"{len(df):,} rows uploaded."
                        st.rerun()
                    except Exception as e:
                        st.session_state.step2_done = False
                        st.session_state.step2_msg  = str(e)
                        st.rerun()

        # Step 3
        step3_disabled = not st.session_state.step2_done
        if st.session_state.step3_done is None:
            step_card(3, "Duplicate Check", "idle", f"Check EIN against `{TBL_IMEIS}`.")
        elif st.session_state.step3_done:
            dup_count = len(st.session_state.duplicates_df) if st.session_state.duplicates_df is not None else 0
            if dup_count > 0:
                step_card(3, "Duplicate Check", "warning", f"{dup_count} duplicate IMEIs found.")
            else:
                step_card(3, "Duplicate Check", "success", "No duplicates. All clear!")
        else:
            step_card(3, "Duplicate Check", "error", st.session_state.step3_msg)

        if not step3_disabled and st.session_state.step3_done is None:
            if st.button("▶ Run Duplicate Check", width="stretch", key="btn3"):
                with st.spinner("Checking..."):
                    try:
                        client = get_bq_client()
                        query = f"""
                            SELECT s.*, CASE WHEN m.EIN IS NOT NULL THEN 'DUPLICATE' ELSE NULL END AS RESULT
                            FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` s
                            LEFT JOIN `{BQ_PROJECT}.{BQ_DATASET}.{TBL_IMEIS}` m ON s.EIN = m.EIN
                        """
                        result_df = client.query(query).to_dataframe()
                        dups_df   = result_df[result_df["RESULT"].notna()]
                        st.session_state.step3_done    = True
                        st.session_state.duplicates_df = dups_df
                        st.rerun()
                    except Exception as e:
                        st.session_state.step3_done = False
                        st.session_state.step3_msg  = str(e)
                        st.rerun()

        if st.session_state.step3_done and st.session_state.duplicates_df is not None:
            dups = st.session_state.duplicates_df
            if len(dups) > 0:
                with st.expander(f"⚠ View {len(dups)} Duplicate Rows"):
                    st.dataframe(dups, width="stretch", hide_index=True)

        # Step 4
        step4_disabled = not st.session_state.step3_done
        if st.session_state.step4_done is None:
            step_card(4, "Push to Master", "idle", f"INSERT into `{TBL_MASTER}`.")
        elif st.session_state.step4_done:
            step_card(4, "Push to Master", "success", st.session_state.step4_msg, f"{BQ_DATASET}.{TBL_MASTER}")
        else:
            step_card(4, "Push to Master", "error", st.session_state.step4_msg)

        if not step4_disabled and st.session_state.step4_done is None:
            if st.button("▶ Push to Master Table", width="stretch", key="btn4"):
                with st.spinner("Pushing..."):
                    try:
                        client = get_bq_client()
                        query = f"""
                            INSERT INTO `{BQ_PROJECT}.{BQ_DATASET}.{TBL_MASTER}`
                                (SO_NO, BILLING_NO, MATL, EIN, DUAL_IMEI_NO, PO_NO, WEEK_NO, PurchaseDate)
                            SELECT SO_NO, BILLING_NO, MATL, EIN, DUAL_IMEI_NO, PO_NO, WEEK_NO, PurchaseDate
                            FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` s
                            WHERE NOT EXISTS (
                                SELECT 1 FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_IMEIS}` m WHERE m.EIN = s.EIN
                            )
                        """
                        job = client.query(query)
                        job.result()
                        st.session_state.step4_done = True
                        st.session_state.step4_msg  = f"{job.num_dml_affected_rows or 0:,} rows inserted."
                        st.rerun()
                    except Exception as e:
                        st.session_state.step4_done = False
                        st.session_state.step4_msg  = str(e)
                        st.rerun()

        # Step 5
        step5_disabled = not st.session_state.step4_done
        if st.session_state.step5_done is None:
            step_card(5, "Clear Staging", "idle", f"Delete all from `{TBL_STAGING}`.")
        elif st.session_state.step5_done:
            step_card(5, "Clear Staging", "success", st.session_state.step5_msg)
        else:
            step_card(5, "Clear Staging", "error", st.session_state.step5_msg)

        if not step5_disabled and st.session_state.step5_done is None:
            if st.button("▶ Clear Staging Table", width="stretch", key="btn5"):
                with st.spinner("Clearing..."):
                    try:
                        client = get_bq_client()
                        client.query(f"DELETE FROM `{BQ_PROJECT}.{BQ_DATASET}.{TBL_STAGING}` WHERE TRUE").result()
                        st.session_state.step5_done = True
                        st.session_state.step5_msg  = "Staging cleared. Pipeline complete!"
                        st.rerun()
                    except Exception as e:
                        st.session_state.step5_done = False
                        st.session_state.step5_msg  = str(e)
                        st.rerun()

        if st.session_state.step5_done:
            st.success("✅ Pipeline complete! All data pushed to master table.")
            st.balloons()

    st.divider()
    st.caption(f"BigQuery: `{BQ_PROJECT}.{BQ_DATASET}` | Staging: `{TBL_STAGING}` | Master: `{TBL_MASTER}`")


# ──────────────────────────────────────────────────────
# TAB 2: IMEI SEARCH
# ──────────────────────────────────────────────────────
with tab_imei:
    st.markdown("### IMEI Search")
    st.markdown("Search one or multiple IMEIs (one per line) across purchases and sales history.")

    imei_input = st.text_area(
        "Enter IMEI number(s) — one per line",
        placeholder="352550421789270\n350505563825367\n350505563825375",
        height=150,
        key="imei_search",
    )

    if st.button("Search IMEI", width="content", key="imei_btn") and imei_input.strip():
        imeis = [x.strip() for x in imei_input.splitlines() if x.strip()]
        imeis_sql = ", ".join(f"'{i}'" for i in imeis)

        with st.spinner(f"Searching BigQuery for {len(imeis)} IMEI(s)..."):
            try:
                client = get_bq_client()

                purch = client.query(f"""
                    SELECT 'purchases' as source, SO_NO, BILLING_NO, MATL as SKU, EIN as IMEI,
                           DUAL_IMEI_NO, PO_NO, WEEK_NO, PurchaseDate as DATE
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases`
                    WHERE EIN IN ({imeis_sql})
                    LIMIT 1000
                """).to_dataframe()

                sales = client.query(f"""
                    SELECT 'sales' as source, IMEI, SKU, INVOICE_NO, CUSTOMER, CARGO_RELEASE,
                           WEEK_NO, DATE
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.sales`
                    WHERE IMEI IN ({imeis_sql})
                    UNION ALL
                    SELECT 'actual_sales' as source, IMEI, SKU, INVOICE_NO, CUSTOMER, CARGO_RELEASE,
                           WEEK_NO, DATE
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales`
                    WHERE IMEI IN ({imeis_sql})
                    LIMIT 1000
                """).to_dataframe()

                found_purch = set(purch["IMEI"]) if len(purch) > 0 else set()
                found_sales = set(sales["IMEI"]) if len(sales) > 0 else set()
                not_found   = [i for i in imeis if i not in found_purch and i not in found_sales]

                st.markdown(f"""
                <div class="metric-row">
                    <div class="metric-box"><div class="metric-val">{len(imeis)}</div><div class="metric-lbl">Searched</div></div>
                    <div class="metric-box"><div class="metric-val" style="color:#00e676">{len(found_purch)}</div><div class="metric-lbl">In Purchases</div></div>
                    <div class="metric-box"><div class="metric-val" style="color:#00e676">{len(found_sales)}</div><div class="metric-lbl">In Sales</div></div>
                    <div class="metric-box"><div class="metric-val" style="color:{'#ff1744' if not_found else '#00e676'}">{len(not_found)}</div><div class="metric-lbl">Not Found</div></div>
                </div>
                """, unsafe_allow_html=True)

                if len(purch) > 0:
                    st.success(f"Found in **purchases** — {len(purch)} record(s)")
                    st.dataframe(purch, width="stretch", hide_index=True)
                else:
                    st.warning("Not found in purchases.")

                if len(sales) > 0:
                    st.success(f"Found in **sales** — {len(sales)} record(s)")
                    st.dataframe(sales, width="stretch", hide_index=True)
                else:
                    st.warning("Not found in sales.")

                if not_found:
                    with st.expander(f"⚠ {len(not_found)} IMEI(s) not found anywhere"):
                        st.code("\n".join(not_found))

            except Exception as e:
                st.error(f"Error: {e}")


# ──────────────────────────────────────────────────────
# TAB 3: MODEL SEARCH
# ──────────────────────────────────────────────────────
with tab_model:
    st.markdown("### Model / SKU Search")
    st.markdown("Search by model name or SKU code across purchases and sales.")

    model_input = st.text_input("Enter Model or SKU", placeholder="e.g. SM-A145 or A145-SM-BK128DL", key="model_search")

    if st.button("Search Model", width="content", key="model_btn") and model_input.strip():
        model = model_input.strip().upper()
        with st.spinner("Searching BigQuery..."):
            try:
                client = get_bq_client()

                purch = client.query(f"""
                    SELECT SO_NO, BILLING_NO, MATL as SKU, EIN as IMEI,
                           PO_NO, WEEK_NO, PurchaseDate
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.purchases`
                    WHERE UPPER(MATL) LIKE '%{model}%'
                    ORDER BY WEEK_NO DESC
                    LIMIT 200
                """).to_dataframe()

                sales = client.query(f"""
                    SELECT 'sales' as source, IMEI, SKU, CUSTOMER, INVOICE_NO, WEEK_NO, DATE
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.sales`
                    WHERE UPPER(SKU) LIKE '%{model}%'
                    UNION ALL
                    SELECT 'actual_sales' as source, IMEI, SKU, CUSTOMER, INVOICE_NO, WEEK_NO, DATE
                    FROM `{BQ_PROJECT}.{BQ_DATASET}.actual_sales`
                    WHERE UPPER(SKU) LIKE '%{model}%'
                    ORDER BY WEEK_NO DESC
                    LIMIT 200
                """).to_dataframe()

                if len(purch) > 0:
                    st.success(f"Found in **purchases** — {len(purch)} record(s)")
                    st.dataframe(purch, width="stretch", hide_index=True)
                else:
                    st.warning("Not found in purchases.")

                if len(sales) > 0:
                    st.success(f"Found in **sales** — {len(sales)} record(s)")
                    st.dataframe(sales, width="stretch", hide_index=True)
                else:
                    st.warning("Not found in sales.")

                if len(purch) == 0 and len(sales) == 0:
                    st.error(f"Model `{model}` not found in any table.")

            except Exception as e:
                st.error(f"Error: {e}")


# ──────────────────────────────────────────────────────
# TAB 4: PROMOTION PDF
# ──────────────────────────────────────────────────────
with tab_promo:
    st.markdown("### Samsung Promotion Bulletin → BigQuery")
    st.markdown("Upload a Samsung 'Sell Thru Promotions Bulletin' PDF. It will be parsed and pushed directly to "
                 "`promotion_data` (normal market) or `promotion_sm` (special market), based on the title.")

    promo_pdf = st.file_uploader("Promotion Bulletin (.pdf)", type=["pdf"], key="promo_pdf")

    if promo_pdf is not None:
        try:
            reference_no, title, target_table, rows, delta_updates, sent_date = parse_promotion_pdf(promo_pdf)

            if not rows and delta_updates:
                target_table, rows, unmatched = resolve_delta_updates(reference_no, delta_updates)
                if not rows:
                    st.error(f"This PDF only lists updated quantities for `{reference_no}`, but none of "
                             f"those models ({', '.join(delta_updates)}) were found in BigQuery for that promotion.")
                    rows = []
                else:
                    title = title or f"Update to promotion {reference_no}"
                    st.info(f"This PDF only lists updated quantities for `{reference_no}`. Matched against "
                            f"existing rows in `{BQ_DATASET}.{target_table}` to fill in product/dates/per-unit.")
                    if unmatched:
                        st.warning(f"Model(s) not found for `{reference_no}` and skipped: {', '.join(unmatched)}")

            st.markdown(f"""
            **Reference Number:** `{reference_no}`
            **Promotion Title:** {title}
            **Date Sent:** {sent_date.strftime('%Y-%m-%d') if sent_date else '_not found in PDF_'}
            **Target table:** `{BQ_DATASET}.{target_table}`
            **Models found:** {len(rows)}
            """)

            if rows:
                preview_df = pd.DataFrame([{
                    "Product": r["PRODUCT"],
                    "Model Code": r["MODEL_CODE"],
                    "Start": r["START_DATE"].strftime("%Y-%m-%d"),
                    "End": r["END_DATE"].strftime("%Y-%m-%d"),
                    "Per Unit": r["PER_UNIT"],
                    "Target Qty": r["QTY"],
                } for r in rows])
                st.dataframe(preview_df, width="stretch", hide_index=True)

                to_insert, duplicates, new_versions = check_promotion_rows(reference_no, target_table, rows)

                if new_versions:
                    st.markdown("#### Updated quantities found — both old and new will be kept in BigQuery")
                    diff_df = pd.DataFrame([{
                        "Model Code": r["MODEL_CODE"],
                        "Previous Qty": prior["QTY_ON_PROMOTION"],
                        "Previous Total Value": prior["TOTAL_VALUE"],
                        "New Qty (this PDF)": r["QTY"],
                        "New Total Value": round(r["PER_UNIT"] * r["QTY"], 2),
                    } for r, prior in new_versions])
                    st.dataframe(diff_df, width="stretch", hide_index=True)

                if duplicates:
                    st.markdown(f"#### {len(duplicates)} duplicate model(s) — already in BigQuery with the same quantity, will be skipped")
                    dup_df = pd.DataFrame([{
                        "Model Code": r["MODEL_CODE"],
                        "Start": r["START_DATE"].strftime("%Y-%m-%d"),
                        "End": r["END_DATE"].strftime("%Y-%m-%d"),
                        "Per Unit": r["PER_UNIT"],
                        "Target Qty": r["QTY"],
                    } for r in duplicates])
                    st.dataframe(dup_df, width="stretch", hide_index=True)

                if st.button("Push to BigQuery", width="stretch", key="promo_push"):
                    with st.spinner("Pushing to BigQuery..."):
                        try:
                            inserted, upload_id = push_promotion_rows(reference_no, title, target_table, to_insert, new_versions, sent_date=sent_date)
                            if inserted:
                                st.success(f"Inserted {inserted} new row(s) into `{BQ_DATASET}.{target_table}` "
                                            f"(previous rows for any updated models were kept as history).")
                                st.balloons()
                                st.session_state["promo_last_push"] = {
                                    "reference_no": reference_no,
                                    "title": title,
                                    "target_table": target_table,
                                    "to_insert": to_insert,
                                    "sent_date": sent_date,
                                }
                            else:
                                st.warning("Nothing new to insert — all models already in BigQuery with the same values.")
                        except Exception as e:
                            st.error(f"Error: {e}")

        except Exception as e:
            st.error(f"Could not parse PDF: {e}")

    # Post-push: extension / price-change detection + SKU-coverage report, and an
    # explicit, separate step to sync the inserted rows into the BACKWORKING sheet.
    last_push = st.session_state.get("promo_last_push")
    if last_push:
        st.markdown("---")
        st.markdown("### Sync to BACKWORKING sheet")
        try:
            per_row_flags, coverage_gaps = analyze_promo_changes(
                last_push["reference_no"], last_push["target_table"], last_push["to_insert"]
            )

            ext = [f for f in per_row_flags if f["extension"]]
            chg = [f for f in per_row_flags if f["price_change"]]

            if ext:
                st.markdown("**Promotion extensions detected** (highlighted blue in the sheet):")
                st.dataframe(pd.DataFrame([{
                    "SKU": f["MODEL_CODE"],
                    "Prior Promo No.": f["prior_promo_no"],
                    "Prior Promo Active": f"{f['prior_start']} → {f['prior_end']}",
                } for f in ext]), width="stretch", hide_index=True)

            if chg:
                st.markdown("**Price changes detected vs. prior promo** (highlighted salmon in the sheet):")
                st.dataframe(pd.DataFrame([{
                    "SKU": f["MODEL_CODE"],
                    "Old Price": f["old_price"],
                    "New Price": f["new_price"],
                } for f in chg]), width="stretch", hide_index=True)

            if not ext and not chg:
                st.info("No extensions or price changes detected vs. promotion history.")

            if coverage_gaps:
                st.markdown("**SKU coverage gaps** — other SKUs in the same model family "
                             "(per the MODELS sheet) that this promo doesn't mention:")
                for gap in coverage_gaps:
                    st.dataframe(pd.DataFrame([{
                        "Sage Code": f["sage_code"],
                        "Samsung Code": f["samsung_code"],
                        "Model": f["model"],
                    } for f in gap["missing_skus"]]), width="stretch", hide_index=True)
                st.caption("These SKUs are not added automatically (no promo quantity exists for them "
                           "in this upload) — review with the team before adding rows manually.")

            if st.button("Sync these rows to the BACKWORKING sheet (with highlights)",
                          width="stretch", key="promo_sheet_sync"):
                with st.spinner("Writing to Google Sheet..."):
                    try:
                        n = sync_promo_rows_to_sheet(
                            last_push["reference_no"], last_push["title"], last_push["target_table"],
                            last_push["to_insert"], per_row_flags, sent_date=last_push.get("sent_date")
                        )
                        st.success(f"Added {n} row(s) to the sheet.")
                        del st.session_state["promo_last_push"]
                    except Exception as e:
                        st.error(f"Error: {e}")
        except Exception as e:
            st.error(f"Could not analyze promotion history: {e}")

# ──────────────────────────────────────────────────────
# Logs tab
# ──────────────────────────────────────────────────────
with tab_logs:
    st.markdown("### Promotion PDF push history")
    st.markdown("Every push from the Promotion PDF tab is logged here. You can revert any push — "
                 "this deletes the row(s) it added and restores the table to how it was before.")

    if st.button("Refresh", key="logs_refresh"):
        st.rerun()

    try:
        logs_df = get_promotion_logs()
    except Exception as e:
        st.error(f"Could not load logs: {e}")
        logs_df = pd.DataFrame()

    if logs_df.empty:
        st.info("No promotion pushes logged yet.")
    else:
        # Group by upload (one push can add multiple rows)
        for upload_id, group in logs_df.groupby("UPLOAD_ID", sort=False):
            first = group.iloc[0]
            ts = first["TIMESTAMP"]
            reverted = bool(first["REVERTED"])
            status = "🔁 Reverted" if reverted else "✅ Active"

            with st.container(border=True):
                st.markdown(f"**{ts}** — `{first['TARGET_TABLE']}` — Promotion `{first['PROMOTION_NO']}` "
                             f"({first['PROMOTION_TITLE']}) — {status}")
                show_df = group[["SAMSUNG_CODE", "ACTION", "OLD_QTY", "NEW_QTY",
                                  "OLD_TOTAL_VALUE", "NEW_TOTAL_VALUE"]].rename(columns={
                    "SAMSUNG_CODE": "Model Code",
                    "ACTION": "Action",
                    "OLD_QTY": "Old Qty",
                    "NEW_QTY": "New Qty",
                    "OLD_TOTAL_VALUE": "Old Total Value",
                    "NEW_TOTAL_VALUE": "New Total Value",
                })
                st.dataframe(show_df, width="stretch", hide_index=True)

                if not reverted:
                    if st.button("Revert this push", key=f"revert_{upload_id}"):
                        with st.spinner("Reverting..."):
                            try:
                                revert_promotion_upload(upload_id, first["TARGET_TABLE"])
                                st.success("Reverted. Refreshing...")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error reverting: {e}")
                else:
                    st.caption(f"Reverted at {first['REVERTED_AT']}")

# ──────────────────────────────────────────────────────
# TAB 6: FINAL SUMMARY (live stock + IMEI + promotion)
# ──────────────────────────────────────────────────────
with tab_summary:
    st.markdown("### Final Summary — Live Stock & Promotions")
    st.markdown("Computed live from `purchases`, `sales`/`actual_sales`, `opening_balance`, "
                 "`promotion_data`/`promotion_sm`, and `report_models` — colored to match the "
                 "FINAL SUMMARY layout from the working sheet.")

    weeks = get_summary_weeks()
    if not weeks:
        st.warning("No WEEK_NO data found in purchases/sales.")
    else:
        col1, col2 = st.columns([1, 3])
        with col1:
            selected_week = st.selectbox("Week", weeks, index=0, key="summary_week")
        with col2:
            st.write("")
            if st.button("🔄 Refresh", key="summary_refresh"):
                build_final_summary.clear()
                get_summary_weeks.clear()
                get_model_code_map.clear()
                get_sam_inventory_overrides.clear()
                st.rerun()

        with st.spinner("Building summary..."):
            df_summary = build_final_summary(selected_week)

        if df_summary.empty:
            st.info(f"No purchase/sale/promotion activity found for `{selected_week}`.")
        else:
            active_promos = (df_summary["PROMO STATUS"] == "ACTIVE").sum()
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-box"><div class="metric-val">{len(df_summary):,}</div><div class="metric-lbl">Models</div></div>
                <div class="metric-box"><div class="metric-val" style="color:#00e676">{active_promos:,}</div><div class="metric-lbl">Active Promotions</div></div>
                <div class="metric-box"><div class="metric-val">{df_summary['IMEI COUNT'].sum():,}</div><div class="metric-lbl">Total IMEIs</div></div>
            </div>
            """, unsafe_allow_html=True)

            base_cols = ["MODEL", "TYPE", "INHAND", "AVAIL. FOR SALE", "ACTUAL SALE", "BALANCE",
                          "SHOW SAM INVENTORY", "WOS"]
            rm_cols = [c for c in df_summary.columns if not c.startswith("SM ")]
            sm_df = df_summary[base_cols + list(SM_RENAME_MAP.keys())].rename(columns=SM_RENAME_MAP)

            col_sync, col_check = st.columns(2)
            with col_sync:
                if st.button("📤 Sync to Google Sheet (BACKWORKING)", key="sync_to_sheet"):
                    try:
                        with st.spinner("Pushing to Google Sheet..."):
                            push_df_to_sheet(df_summary[rm_cols], RM_SYNC_TAB)
                            push_df_to_sheet(sm_df, SM_SYNC_TAB)
                        st.success("Synced FINAL SUMMARY and SM - FINAL SUMMARY tabs to the BACKWORKING sheet.")
                    except Exception as e:
                        st.error(f"Sync failed: {e}")

            with col_check:
                if st.button("🔍 Check Sheet for Manual Edits", key="check_sheet_edits"):
                    try:
                        with st.spinner("Comparing sheet to BigQuery..."):
                            diffs = detect_sheet_edits(selected_week, df_summary[rm_cols], sm_df)
                            n_new = record_pending_edits(selected_week, diffs)
                        if diffs:
                            st.warning(f"Found {len(diffs)} changed cell(s) in the sheet ({n_new} new).")
                        else:
                            st.success("No manual edits detected — sheet matches BigQuery.")
                    except Exception as e:
                        st.error(f"Check failed: {e}")

            pending = get_pending_sheet_edits(selected_week)
            if not pending.empty:
                st.markdown(f"#### ⏳ Pending Approval — {len(pending)} change(s) made directly in the Google Sheet")
                st.markdown("These cells differ from the BigQuery-computed values. Approve to apply them as "
                             "overrides (kept on top of BigQuery going forward), or reject to let the next "
                             "sync overwrite them back to the computed value.")
                for _, edit in pending.iterrows():
                    c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
                    with c1:
                        st.markdown(f"**{edit['model']}** — `{edit['column_name']}` "
                                     f"({edit['tab_name'].replace(' (App Sync)','')}): "
                                     f"`{edit['computed_value']}` → `{edit['sheet_value']}`")
                    with c2:
                        if st.button("✅ Approve", key=f"approve_{edit['edit_id']}"):
                            resolve_pending_edit(edit['edit_id'], True, selected_week,
                                                  model=edit['model'], tab_name=edit['tab_name'],
                                                  column=edit['column_name'],
                                                  value=edit['sheet_value'])
                            build_final_summary.clear()
                            st.rerun()
                    with c3:
                        if st.button("❌ Reject", key=f"reject_{edit['edit_id']}"):
                            resolve_pending_edit(edit['edit_id'], False, selected_week)
                            st.rerun()

            sub_rm, sub_sm = st.tabs(["Final Summary (RM)", "SM - Final Summary"])

            with sub_rm:
                st.dataframe(style_final_summary(df_summary[rm_cols]), width="stretch", hide_index=True)

            with sub_sm:
                st.dataframe(style_final_summary(sm_df), width="stretch", hide_index=True)

            flagged = df_summary[df_summary["ADJUSTED CHECK"] != ""]
            if not flagged.empty:
                st.markdown(f"#### ⚠️ Needs Check ({len(flagged)} model{'s' if len(flagged) != 1 else ''})")
                st.markdown("These models have a manually entered **SHOW SAM INVENTORY** that differs "
                             "from our computed **INHAND** — review the **ADJUSTED** figure.")
                st.dataframe(
                    style_final_summary(flagged[["MODEL", "TYPE", "INHAND", "SHOW SAM INVENTORY",
                                                   "AVAIL. FOR SALE", "ADJUSTED", "ADJUSTED CHECK"]]),
                    width="stretch", hide_index=True,
                )

            with st.expander("✏️ Edit SHOW SAM INVENTORY (manual override)"):
                st.markdown("By default **SHOW SAM INVENTORY** equals our computed **INHAND**. "
                             "Enter a value here only when Samsung's reported on-hand for a model "
                             "differs from ours — this will recompute **ADJUSTED** and raise the "
                             "**ADJUSTED CHECK** flag for that model.")
                override_models = df_summary["MODEL"].tolist()
                sel_model = st.selectbox("Model", override_models, key="sam_override_model")
                current_row = df_summary[df_summary["MODEL"] == sel_model].iloc[0]
                new_val = st.number_input(
                    "SHOW SAM INVENTORY",
                    min_value=0,
                    value=int(current_row["SHOW SAM INVENTORY"]),
                    step=1,
                    key="sam_override_value",
                )
                if st.button("Save override", key="sam_override_save"):
                    save_sam_inventory_override(selected_week, sel_model, new_val)
                    get_sam_inventory_overrides.clear()
                    build_final_summary.clear()
                    st.success(f"Saved SHOW SAM INVENTORY = {new_val} for {sel_model} ({selected_week}).")
                    st.rerun()

# ──────────────────────────────────────────────────────
# TAB 7: PROMOTIONS — live colored frame (Sell-Thru + Special Market)
# ──────────────────────────────────────────────────────
with tab_promotions:
    st.markdown("### Promotions — Live Frame")
    st.markdown("Computed live from `promotion_data` (Sell-Thru) and `promotion_sm` (Special Market), "
                 "colored to match the Final Summary palette. ACTIVE/EXPIRED is based on END DATE vs today.")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("🔄 Refresh", key="promotions_refresh"):
            build_promotions_frame.clear()
            get_model_code_map.clear()
            st.rerun()
    with col_b:
        if st.button("📤 Append new promos to BACKWORKING sheet", key="append_promos_sheet"):
            try:
                with st.spinner("Checking for new promo rows..."):
                    n1, skip1 = sync_promotion_table_append("promotion_data", "PROMOTION DATA")
                    n2, skip2 = sync_promotion_table_append("promotion_sm", "PROMOTION - SM")
                st.success(f"Appended {n1} new row(s) to PROMOTION DATA, {n2} new row(s) to PROMOTION - SM. "
                           "Existing rows/formulas were not touched.")
                if skip1 or skip2:
                    st.warning(
                        f"Skipped {skip1} row(s) and {skip2} row(s) due to unknown errors."
                    )
            except Exception as e:
                import traceback
                st.error(f"Append failed: {e}")
                st.error(f"Details: {traceback.format_exc()}")

    with st.spinner("Building promotions frame..."):
        df_promos = build_promotions_frame()

    if df_promos.empty:
        st.info("No promotion records found.")
    else:
        view = st.radio(
            "View",
            ["All Promotions (List)", "SKU Detail"],
            horizontal=True,
            key="promotions_view_mode",
        )
        active_count = (df_promos["STATUS"] == "ACTIVE").sum()

        if view == "All Promotions (List)":
            df_list = build_promotions_list(df_promos)
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-box"><div class="metric-val">{len(df_list):,}</div><div class="metric-lbl">Total Promotions</div></div>
                <div class="metric-box"><div class="metric-val" style="color:#00e676">{(df_list['STATUS']=='ACTIVE').sum():,}</div><div class="metric-lbl">Active</div></div>
                <div class="metric-box"><div class="metric-val">{(df_list['STATUS']=='EXPIRED').sum():,}</div><div class="metric-lbl">Expired</div></div>
            </div>
            """, unsafe_allow_html=True)
            search = st.text_input("Filter by Promo Number", key="promo_list_search", placeholder="e.g. A0S3A0A260000849")
            if search:
                df_list = df_list[df_list["PROMO NUMBER"].str.contains(search, case=False, na=False)]
            st.caption(f"Every promotion ever pushed to `promotion_data` / `promotion_sm`, one row per Promo Number — {len(df_list):,} shown.")
            st.dataframe(style_promotions(df_list), width="stretch", hide_index=True)
        else:
            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-box"><div class="metric-val">{len(df_promos):,}</div><div class="metric-lbl">Total Promotions</div></div>
                <div class="metric-box"><div class="metric-val" style="color:#00e676">{active_count:,}</div><div class="metric-lbl">Active</div></div>
                <div class="metric-box"><div class="metric-val">{(df_promos['STATUS']=='EXPIRED').sum():,}</div><div class="metric-lbl">Expired</div></div>
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(style_promotions(df_promos), width="stretch", hide_index=True)

# ──────────────────────────────────────────────────────
# TAB 8: CLOSING — IMEI reconciliation (in vs out, orphans, PO traceability)
# ──────────────────────────────────────────────────────
with tab_closing:
    st.markdown("### Closing — IMEI Reconciliation")
    st.markdown("IMEI in = out check across `purchases`, `sales`/`actual_sales`: total balance, "
                 "orphan sold IMEIs (sales with no purchase record), and "
                 "PO-by-PO IMEI traceability (which PO's IMEIs went to which sale).")

    if st.button("🔄 Refresh", key="closing_refresh"):
        build_closing_overview.clear()
        get_orphan_imeis.clear()
        get_orphan_summary_by_sku.clear()
        build_po_traceability.clear()
        get_po_imei_detail.clear()
        st.rerun()

    with st.spinner("Computing overview..."):
        ov = build_closing_overview()

    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-box"><div class="metric-val">{ov['total_received']:,}</div><div class="metric-lbl">Total Received</div></div>
        <div class="metric-box"><div class="metric-val" style="color:#00e676">{ov['imeis_sold']:,}</div><div class="metric-lbl">Sold</div></div>
        <div class="metric-box"><div class="metric-val" style="color:#00e5ff">{ov['imeis_in_hand']:,}</div><div class="metric-lbl">In Hand</div></div>
        <div class="metric-box"><div class="metric-val" style="color:#ffab00">{ov['orphan_imeis']:,}</div><div class="metric-lbl">Orphan Sold IMEIs (no purchase)</div></div>
        <div class="metric-box"><div class="metric-val">{ov['duplicate_imeis']:,}</div><div class="metric-lbl">Duplicate IMEIs</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    sub_trace, sub_orphan, sub_po = st.tabs(
        ["Search IMEI", "Orphan Sold IMEIs", "PO-by-PO Traceability"])

    # --- Search a single IMEI: full in/out trace ---
    with sub_trace:
        st.markdown("Search any IMEI to see its **full in/out trace** — which PO it was received on, "
                     "the purchase date, and (if sold) the invoice/customer/date it went out on. "
                     "Works the same way as searching Samsung Weekly / Magaya sheets, but live from the database.")
        imei_q = st.text_input("IMEI", key="closing_imei_search", placeholder="e.g. 355875111442113")
        if imei_q:
            with st.spinner("Tracing IMEI..."):
                trace_df = get_imei_trace(imei_q)
            if trace_df.empty:
                st.warning("IMEI not found in purchases or sales.")
            else:
                row = trace_df.iloc[0]
                st.markdown(f"**Status:** {row['STATUS']}")
                st.dataframe(trace_df, width="stretch", hide_index=True)

    # --- Orphan sold IMEIs ---
    with sub_orphan:
        st.markdown("IMEIs that appear in `sales`/`actual_sales` but have **no matching record in `purchases`** — "
                     "i.e. sold but never recorded as received.")

        with st.spinner("Loading SKU summary..."):
            summary_df = get_orphan_summary_by_sku()
        st.dataframe(summary_df, width="stretch", hide_index=True, height=250)

        sku_search = st.text_input("Filter by SKU/model (e.g. A315, R177)", key="orphan_sku_search")
        with st.spinner("Loading orphan IMEIs..."):
            orphan_df = get_orphan_imeis(sku_search)
        st.caption(f"Showing {len(orphan_df):,} rows (max 5,000)")
        st.dataframe(orphan_df, width="stretch", hide_index=True, height=350)
        if not orphan_df.empty:
            st.download_button("⬇ Download orphan IMEIs (CSV)", orphan_df.to_csv(index=False),
                                file_name="orphan_sold_imeis_no_purchase_record.csv", mime="text/csv")

    # --- PO-by-PO traceability ---
    with sub_po:
        st.markdown("For each PO: IMEIs received, how many were sold vs still in hand. "
                     "Select a PO to drill into the IMEI-level in/out detail (which sale/invoice each IMEI went to).")
        with st.spinner("Loading PO summary..."):
            po_df = build_po_traceability()

        po_search = st.text_input("Filter by PO_NO or Model code", key="po_search")
        view_df = po_df
        if po_search:
            mask = po_df["PO_NO"].astype(str).str.contains(po_search, case=False, na=False) | \
                   po_df["sample_matl"].astype(str).str.contains(po_search, case=False, na=False)
            view_df = po_df[mask]
        st.caption(f"Showing {len(view_df):,} of {len(po_df):,} POs")
        st.dataframe(view_df, width="stretch", hide_index=True, height=350)
        if not po_df.empty:
            st.download_button("⬇ Download PO-by-PO traceability (CSV)", po_df.to_csv(index=False),
                                file_name="po_by_po_imei_traceability.csv", mime="text/csv")

        st.markdown("#### Drill into a PO")
        po_options = po_df["PO_NO"].astype(str).tolist()
        selected_po = st.selectbox("PO_NO", po_options, key="po_drill_select")
        if selected_po:
            with st.spinner(f"Loading IMEI detail for PO {selected_po}..."):
                detail_df = get_po_imei_detail(selected_po)
            st.caption(f"{len(detail_df):,} IMEIs (max 5,000 shown)")
            st.dataframe(detail_df, width="stretch", hide_index=True, height=350)
            if not detail_df.empty:
                st.download_button("⬇ Download this PO's IMEI detail (CSV)", detail_df.to_csv(index=False),
                                    file_name=f"po_{selected_po}_imei_detail.csv", mime="text/csv")

# ──────────────────────────────────────────────────────
# TAB 8: AI CHAT (agentic, full data access)
# ──────────────────────────────────────────────────────
with tab_chat:
    st.markdown("### Chat with your data")
    st.markdown(f"Ask anything about purchases, sales, IMEIs, models, or promotions. The assistant has live "
                 f"read access to every table in `{BQ_DATASET}` via BigQuery — it always queries the real data "
                 f"instead of guessing.")

    st.caption("This assistant can also make data changes (INSERT/UPDATE/DELETE) when you ask — "
               "it will always show you the exact SQL and ask for confirmation first. Every change is "
               "logged to `ai_chat_write_logs`.")

    for key, default in [("chat_display", []), ("chat_anthropic_history", []), ("chat_pending", None), ("chat_dfs", [])]:
        if key not in st.session_state:
            st.session_state[key] = default

    has_key = bool(safe_secret("ANTHROPIC_API_KEY")) or bool(__import__("os").environ.get("ANTHROPIC_API_KEY"))
    if not has_key:
        st.warning("No `ANTHROPIC_API_KEY` found in secrets or environment. Add it to enable the chat.")

    for i, msg in enumerate(st.session_state.chat_display):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            df = st.session_state.chat_dfs[i] if i < len(st.session_state.chat_dfs) else None
            if df is not None and not df.empty:
                st.download_button(
                    f"📥 Download full result ({len(df):,} rows) as CSV",
                    df.to_csv(index=False).encode("utf-8"),
                    file_name="query_result.csv",
                    mime="text/csv",
                    key=f"dl_{i}",
                )

    def advance_agent():
        """Run the agent loop from current history; updates session state and re-renders."""
        st.session_state.pop("_chat_pending_df", None)
        st.session_state.pop("_chat_pending_sql", None)
        while True:
            status, payload, messages = run_chat_agent(st.session_state.chat_anthropic_history)
            st.session_state.chat_anthropic_history = messages
            if status == "final":
                st.session_state.chat_display.append({"role": "assistant", "content": payload})
                st.session_state.chat_dfs.append(st.session_state.pop("_chat_pending_df", None))
                return
            # pending_write
            st.session_state.chat_pending = payload
            return

    if st.session_state.chat_pending is None:
        if prompt := st.chat_input("Ask about purchases, sales, IMEIs, models, promotions...", disabled=not has_key):
            st.session_state.chat_display.append({"role": "user", "content": prompt})
            st.session_state.chat_dfs.append(None)
            st.session_state.chat_anthropic_history.append({"role": "user", "content": prompt})

            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        advance_agent()
                    except Exception as e:
                        st.error(f"Error: {e}")
            st.rerun()
    else:
        write_block = st.session_state.chat_pending["write_block"]
        sql_text = write_block.input.get("query", "")
        with st.chat_message("assistant"):
            st.markdown("I'd like to run the following data change. Confirm to proceed:")
            st.code(sql_text, language="sql")
            c1, c2 = st.columns(2)
            confirm = c1.button("✅ Confirm & Run", width="stretch", key="confirm_write")
            cancel  = c2.button("❌ Cancel", width="stretch", key="cancel_write")

            if confirm or cancel:
                pending = st.session_state.chat_pending
                tool_results = list(pending["pending_results"])
                if confirm:
                    with st.spinner("Running..."):
                        result = _bq_tool_run_write_sql(sql_text)
                else:
                    result = {"status": "cancelled", "message": "User declined this write operation."}
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": write_block.id,
                    "content": json.dumps(result, default=str),
                })
                st.session_state.chat_anthropic_history.append({"role": "user", "content": tool_results})
                st.session_state.chat_pending = None
                with st.spinner("Thinking..."):
                    try:
                        advance_agent()
                    except Exception as e:
                        st.error(f"Error: {e}")
                st.rerun()

    if st.session_state.chat_display:
        if st.button("Clear chat", key="clear_chat"):
            st.session_state.chat_display = []
            st.session_state.chat_anthropic_history = []
            st.session_state.chat_pending = None
            st.session_state.chat_dfs = []
            st.rerun()
