## Marketing Channel Reporting: Data Cleanup Pipeline

### Problem
Sales, web analytics, and channel spend data from three systems with inconsistent naming, mixed date formats, and duplicates.

### Solution
Python pipeline that:
- Standardizes channel names across three sources
- Normalizes date formats
- Dedupes on order_id
- Corrects negative sessions and strips `$` and commas from text spend values
- Keeps missing spend blank instead of zero, so ROAS is not distorted
- Merges the three sources into one month x channel table
- Reconciles against the generated ground truth

### Files
- `generate.py` — generates synthetic three-source dataset into `original_data/`
- `clean.py` — main cleaning pipeline
- `original_data/` — messy input files (`crm_orders.csv`, `web_sessions.csv`, `channel_spend.xlsx`)
- `cleaned_data/` — cleaned output as `.csv` and `.xlsx`; `channel_performance.csv` is the merged table

### Result
One validated, analysis-ready table (`channel_performance.csv`) of spend, revenue, orders, sessions, and bounced sessions by channel and month, ready for revenue, spend, and ROAS reporting in a BI tool such as Tableau. Spend, orders, and sessions match the ground truth exactly; revenue is within 0.05%, depending on which copy of a duplicated order is kept.

### To run
```bash
pip install pandas numpy openpyxl
python generate.py  # Creates original_data/*
python clean.py     # Outputs cleaned_data/*
```

Data is synthetic, generated to mirror real multi-system reporting scenarios.
