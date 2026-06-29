import gspread
from google.oauth2 import service_account
import toml

creds = toml.load('.streamlit/secrets.toml')['gcp_service_account']
gc = gspread.service_account_from_dict(creds)
sh = gc.open_by_key('1gHeuIVhf64PTMEU_fE0puZTJR9ejZ6cQ2gZpSJx2cRo')

targets = [
    'IMEI_SAM_15_DAYS_WORK', 
    'IMEI_SAM_15_DAYS', 
    'IMEI SOLD', 
    'IMEI SALE SUMMARY',
    'DATA',
    'Stock'
]

for ws in sh.worksheets():
    if ws.title in targets:
        try:
            vals = ws.get_all_values()
            last_filled = max([i for i, row in enumerate(vals) if any(row)]) + 1 if vals and any(any(row) for row in vals) else 1
            # Keep a buffer of 10 rows
            new_size = max(last_filled + 10, 100) 
            
            if ws.row_count > new_size:
                print(f"Resizing {ws.title} from {ws.row_count} to {new_size} rows.")
                ws.resize(rows=new_size)
        except Exception as e:
            print(f"Error processing {ws.title}: {e}")

print(f"New total cells: {sum([ws.row_count * ws.col_count for ws in sh.worksheets()])}")
