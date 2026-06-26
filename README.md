# Samsung Purchases Pipeline

Streamlit app for processing Samsung invoice files and pushing data to BigQuery.

## Setup
```bash
pip install -r requirements.txt
gcloud auth application-default login
streamlit run app.py
```

## BigQuery
- Project: `psi-reports-493216`
- Dataset: `sam`
