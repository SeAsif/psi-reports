"""One-off backfill: ingest any PROMOTION DATA / PROMOTION - SM rows present in the
new primary sheet (BACKWORKING_SHEET_ID) that aren't yet in BigQuery promotion_data/
promotion_sm. Run once. Not part of the Streamlit app.

Usage:
  GOOGLE_APPLICATION_CREDENTIALS=/Users/apple/.config/gcloud/application_default_credentials.json \
    /Users/apple/Downloads/samsung_pipeline_venv/bin/python3 backfill_promo_from_new_sheet.py
"""
import re
import uuid
import toml
import gspread
import pandas as pd
from datetime import datetime
from google.cloud import bigquery

BQ_PROJECT = "psi-reports-493216"
BQ_DATASET = "sam"
PRIMARY_SHEET_ID = "1NkEOEjGy9czmaLOf5AgDWQSIgZUVceLxPRosGuj_8_Y"

TARGETS = [
    ("promotion_data", 183516201),   # PROMOTION DATA
    ("promotion_sm", 2031721001),    # PROMOTION - SM
]

KEY_COLS = ["PROMOTION_NO", "SAMSUNG_CODE", "START_DATE", "QTY_ON_PROMOTION"]


def norm_date(v):
    s = str(v or "").strip()
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


def norm_num(v):
    s = str(v or "").strip().replace("$", "").replace(",", "")
    if not s:
        return ""
    try:
        return str(float(s))
    except ValueError:
        return s


def row_key(rec):
    return (
        str(rec.get("PROMOTION_NO", "")).strip(),
        str(rec.get("SAMSUNG_CODE", "")).strip(),
        norm_date(rec.get("START_DATE")),
        norm_num(rec.get("QTY_ON_PROMOTION")),
    )


def main(dry_run=False):
    secrets = toml.load("/Users/apple/Downloads/bigquery/.streamlit/secrets.toml")
    gc = gspread.service_account_from_dict(
        dict(secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    bq = bigquery.Client(project=BQ_PROJECT)
    sh = gc.open_by_key(PRIMARY_SHEET_ID)

    for table_name, gid in TARGETS:
        table_id = f"{BQ_PROJECT}.{BQ_DATASET}.{table_name}"
        table = bq.get_table(table_id)
        all_columns = [f.name for f in table.schema]
        raw_columns = [c for c in all_columns if c not in ("UPLOAD_ID", "SENT_DATE")]

        ws = sh.get_worksheet_by_id(gid)
        values = ws.get_all_values()
        sheet_header = values[1]
        data_rows = values[2:]

        if len(sheet_header) < len(raw_columns):
            raise RuntimeError(
                f"{table_name}: sheet has {len(sheet_header)} cols, expected >= {len(raw_columns)}"
            )

        existing_keys = set()
        existing_q = f"SELECT {', '.join(KEY_COLS)} FROM `{table_id}`"
        for r in bq.query(existing_q):
            existing_keys.add(row_key(dict(r.items())))

        upload_id = str(uuid.uuid4())
        new_records = []
        for raw_row in data_rows:
            if not any(v.strip() for v in raw_row[:12]):
                continue  # fully blank row
            rec = {col: (raw_row[i] if i < len(raw_row) else "") for i, col in enumerate(raw_columns)}
            rec = {k: (v if v != "" else None) for k, v in rec.items()}
            key = row_key(rec)
            if key in existing_keys:
                continue
            rec["UPLOAD_ID"] = upload_id
            rec["SENT_DATE"] = None
            new_records.append(rec)
            existing_keys.add(key)  # guard against dupes within the sheet itself

        print(f"{table_name}: {len(data_rows)} sheet rows, {len(new_records)} new to insert")
        for rec in new_records[:10]:
            print("  +", rec.get("PROMOTION_NO"), rec.get("SAMSUNG_CODE"), rec.get("START_DATE"), rec.get("QTY_ON_PROMOTION"))

        if new_records and not dry_run:
            df = pd.DataFrame(new_records, columns=all_columns)
            job_config = bigquery.LoadJobConfig(write_disposition=bigquery.WriteDisposition.WRITE_APPEND)
            job = bq.load_table_from_dataframe(df, table_id, job_config=job_config)
            job.result()
            print(f"{table_name}: inserted {len(df)} rows (upload_id={upload_id})")


if __name__ == "__main__":
    import sys
    main(dry_run="--dry-run" in sys.argv)
