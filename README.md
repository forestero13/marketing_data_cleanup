## Marketing Channel Reporting: Data Cleanup Pipeline

### Problem
Sales, web analytics, and channel spend data from three systems with inconsistent naming, mixed date formats, and duplicates.

### Solution
Python pipeline that:
- Standardizes channel names across three sources
- Normalizes date formats
- Dedupes on order_id
- Removes negative sessions and malformed spend data
- Validates against generated ground truth

### Files
- `generate.py` — generates synthetic three-source dataset
- `clean.py` — main cleaning pipeline
- `original_data/` — messy input CSVs
- `cleaned_data/` — cleaned output, ready for analysis

### Result
One validated, analysis-ready table fed into Tableau dashboard showing revenue, spend, and ROAS by channel and month.

### To run
```python
python generate.py  # Creates original_data/*.csv
python clean.py     # Outputs cleaned_data/marketing_clean.csv
```

Data is synthetic, generated to mirror real multi-system reporting scenarios.
