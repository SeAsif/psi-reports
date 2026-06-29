import gspread
from google.oauth2 import service_account
import toml

creds = toml.load('.streamlit/secrets.toml')['gcp_service_account']
gc = gspread.service_account_from_dict(creds)
sh = gc.open_by_key('1gHeuIVhf64PTMEU_fE0puZTJR9ejZ6cQ2gZpSJx2cRo')

for ws in sh.worksheets():
    if ws.row_count * ws.col_count > 100000:
        try:
            vals = ws.get_all_values()
            last_filled = max([i for i, row in enumerate(vals) if any(row)]) + 1 if vals else 0
            empty_rows = ws.row_count - last_filled
            print(f'{ws.title}: {empty_rows} empty rows out of {ws.row_count}, freeing {empty_rows * ws.col_count} cells if deleted')
        except Exception as e:
            print(f'{ws.title}: error {e}')
