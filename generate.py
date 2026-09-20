#!/usr/bin/env python3
"""Generate the synthetic crm_orders.csv, web_sessions.csv and channel_spend.xlsx (Jan 2025 - Jun 2026)."""

import calendar

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

# Configuration

SEED = 20250101
rng = np.random.default_rng(SEED)

MONTHS = pd.period_range("2025-01", "2026-06", freq="M")  # 18 months

# Per-channel ground truth: baseline monthly spend (USD), ROAS, average order value, conversion rate, bounce rate
CHANNELS = {
    "Google Ads": dict(spend=4500.0, roas=3.00, aov=290.0, cvr=0.021, bounce=0.47),
    "Meta":       dict(spend=2600.0, roas=2.40, aov=240.0, cvr=0.014, bounce=0.58),
    "Email":      dict(spend=780.0,  roas=5.20, aov=320.0, cvr=0.035, bounce=0.44),
    "Organic":    dict(spend=520.0,  roas=5.80, aov=300.0, cvr=0.028, bounce=0.41),
    "Affiliate":  dict(spend=1200.0, roas=1.90, aov=210.0, cvr=0.018, bounce=0.63),
}

# mild seasonality: Q4 peak, soft Jan/Feb
SPEND_SEASONALITY = {
    1: 0.88, 2: 0.85, 3: 0.95, 4: 0.97, 5: 1.00, 6: 1.00,
    7: 0.95, 8: 0.97, 9: 1.05, 10: 1.15, 11: 1.35, 12: 1.30,
}
# efficiency improves slightly into the holidays
ROAS_SEASONALITY = {
    1: 0.95, 2: 0.95, 3: 1.00, 4: 1.00, 5: 1.00, 6: 1.00,
    7: 0.98, 8: 1.00, 9: 1.02, 10: 1.05, 11: 1.12, 12: 1.08,
}
MONTHLY_GROWTH = 0.004        # ~7% underlying growth across the 18 months
SPEND_NOISE_SIGMA = 0.08      # lognormal
ROAS_NOISE_SIGMA = 0.10       # lognormal

# weekday shape for daily sessions and orders (Mon=0 ... Sun=6)
WEEKDAY_WEIGHTS = {0: 1.06, 1: 1.06, 2: 1.03, 3: 1.00, 4: 0.94, 5: 0.88, 6: 1.03}

# Naming variants: each source system spells the channels differently

CRM_VARIANTS = {
    "Google Ads": ["Google Ads", "google_ads", "GoogleAds", "GOOGLE ADS"],
    "Meta":       ["Meta", "meta_ads", "Facebook", "FB Ads"],
    "Email":      ["Email", "email_marketing", "E-mail", "EMAIL"],
    "Organic":    ["Organic", "organic_search", "ORGANIC", "Organic Search"],
    "Affiliate":  ["Affiliate", "affiliate", "AFFILIATE", "Affiliates"],
}

WEB_VARIANTS = {
    "Google Ads": ["google / cpc", "google-ads", "adwords", "googleads"],
    "Meta":       ["facebook / paid", "fb_paid", "meta / social", "FB"],
    "Email":      ["email / newsletter", "klaviyo", "email-campaign", "Email"],
    "Organic":    ["(organic)", "google / organic", "organic", "seo"],
    "Affiliate":  ["affiliate / partner", "aff", "partner-network", "AFFILIATE"],
}

SPEND_VARIANTS = {
    "Google Ads": ["Paid Search - Google", "paid search google", "PAID SEARCH - GOOGLE"],
    "Meta":       ["Paid Social - Meta", "paid social meta", "PAID SOCIAL - META"],
    "Email":      ["Email Marketing", "email mktg", "EMAIL MARKETING"],
    "Organic":    ["Organic / SEO", "organic seo", "ORGANIC / SEO"],
    "Affiliate":  ["Affiliate Partners", "affiliate partners", "AFFILIATE PARTNERS"],
}

REGIONS = ["Northeast", "Southeast", "Midwest", "West", "Southwest"]
REGION_WEIGHTS = [0.24, 0.21, 0.18, 0.27, 0.10]

N_ORDER_ROWS = 2000           # total rows written to crm_orders.csv
N_DUPLICATE_IDS = 40          # duplicated order_ids (extra rows, slightly different revenue)
N_BASE_ORDERS = N_ORDER_ROWS - N_DUPLICATE_IDS
N_CUSTOMERS = 1300
REGION_NULL_RATE = 0.03

TARGET_SESSION_ROWS = 5000
N_NEGATIVE_SESSION_ROWS = 15
EXTRA_FRAGMENT_P = 0.415      # tuned so mean fragments/day ~1.83 -> ~5,000 rows

SPEND_TEXT_SHARE = 0.40       # share of Spend_USD cells written as "$12,450" strings
N_AFFILIATE_GAPS = 2


# Helpers

def largest_remainder(weights, total):
    """Split an integer total across weights with no rounding drift."""
    w = np.asarray(weights, dtype=float)
    w = w / w.sum() * total
    base = np.floor(w).astype(int)
    short = int(total - base.sum())
    if short > 0:
        base[np.argsort(-(w - base))[:short]] += 1
    return base


def banner(title):
    print("=" * 78)
    print(title)
    print("=" * 78)


def month_days(period):
    """Every day of a pandas Period (month) as a Timestamp."""
    year, month = period.year, period.month
    n = calendar.monthrange(year, month)[1]
    return [pd.Timestamp(year=year, month=month, day=d) for d in range(1, n + 1)]


def day_weights(days):
    """Weekday shape plus noise, normalised."""
    base = np.array([WEEKDAY_WEIGHTS[d.weekday()] for d in days], dtype=float)
    base *= rng.lognormal(0.0, 0.18, size=len(days))
    return base / base.sum()


# 1. Ground-truth monthly model
rows = []
for i, m in enumerate(MONTHS):
    growth = 1.0 + MONTHLY_GROWTH * i
    for ch, p in CHANNELS.items():
        spend = (p["spend"] * SPEND_SEASONALITY[m.month] * growth
                 * rng.lognormal(0.0, SPEND_NOISE_SIGMA))
        roas = p["roas"] * ROAS_SEASONALITY[m.month] * rng.lognormal(0.0, ROAS_NOISE_SIGMA)
        rows.append(dict(month=m, channel=ch, model_spend=spend, model_revenue=spend * roas))

truth = pd.DataFrame(rows)

# orders = revenue / AOV, forced to the exact row count
truth["raw_orders"] = truth.apply(
    lambda r: r["model_revenue"] / CHANNELS[r["channel"]]["aov"], axis=1
)
truth["orders"] = largest_remainder(truth["raw_orders"].values, N_BASE_ORDERS)
truth.loc[truth["orders"] < 1, "orders"] = 1

# sessions = orders / CVR, so conversion holds exactly after cleaning
truth["sessions"] = truth.apply(
    lambda r: int(round(r["orders"] / CHANNELS[r["channel"]]["cvr"])), axis=1
)


# 2. crm_orders.csv
customer_pool = np.array([f"CUST-{i:05d}" for i in range(10000, 10000 + N_CUSTOMERS)])
# skewed customer frequency: a minority place most repeat orders
customer_p = rng.lognormal(0.0, 0.7, size=N_CUSTOMERS)
customer_p /= customer_p.sum()

order_rows = []
for rec in truth.itertuples():
    n = int(rec.orders)

    # split the month's revenue across its orders exactly
    w = rng.lognormal(0.0, 0.55, size=n)
    values = np.round(rec.model_revenue * w / w.sum(), 2)

    days = month_days(rec.month)
    picked = rng.choice(len(days), size=n, p=day_weights(days))

    order_rows.append(pd.DataFrame({
        "true_channel": rec.channel,
        "month": str(rec.month),
        "order_date": [days[j] for j in picked],
        "revenue": values,
        "customer_id": rng.choice(customer_pool, size=n, p=customer_p),
    }))

orders = pd.concat(order_rows, ignore_index=True).sort_values("order_date").reset_index(drop=True)
orders["order_id"] = [f"ORD-{100000 + i}" for i in range(len(orders))]
orders["channel"] = [rng.choice(CRM_VARIANTS[c]) for c in orders["true_channel"]]
orders["region"] = rng.choice(REGIONS, size=len(orders), p=REGION_WEIGHTS)
orders.loc[rng.random(len(orders)) < REGION_NULL_RATE, "region"] = np.nan

# ground-truth revenue comes from the de-duplicated rows only
true_revenue = orders.groupby(["month", "true_channel"], as_index=False)["revenue"].sum()

# Duplicated order_ids: same id, slightly different revenue, sometimes a different channel spelling or region
dupe_idx = rng.choice(len(orders), size=N_DUPLICATE_IDS, replace=False)
dupes = orders.loc[dupe_idx].copy()
dupes["revenue"] = np.round(dupes["revenue"] * (1 + rng.uniform(-0.06, 0.06, len(dupes))), 2)
dupes["channel"] = [rng.choice(CRM_VARIANTS[c]) for c in dupes["true_channel"]]
flip = rng.random(len(dupes)) < 0.35
dupes.loc[flip, "region"] = rng.choice(REGIONS, size=int(flip.sum()), p=REGION_WEIGHTS)

crm = pd.concat([orders, dupes], ignore_index=True)
crm = crm.sample(frac=1.0, random_state=7).reset_index(drop=True)
crm_out = pd.DataFrame({
    "order_id": crm["order_id"],
    "order_date": crm["order_date"].dt.strftime("%m/%d/%Y"),
    "customer_id": crm["customer_id"],
    "channel": crm["channel"],
    "revenue": crm["revenue"].astype(float),
    "region": crm["region"],
})
crm_out.to_csv("crm_orders.csv", index=False)


# 3. web_sessions.csv
session_rows = []
for rec in truth.itertuples():
    days = month_days(rec.month)
    daily = rng.multinomial(int(rec.sessions), day_weights(days))

    for day, total in zip(days, daily):
        if total <= 0:
            continue
        n_frag = 1 + int(rng.binomial(2, EXTRA_FRAGMENT_P))
        parts = rng.multinomial(int(total), np.ones(n_frag) / n_frag)
        for part in parts:
            if part <= 0:
                continue
            bounce = CHANNELS[rec.channel]["bounce"] + rng.normal(0.0, 0.055)
            session_rows.append((
                day.strftime("%Y-%m-%d"),
                rng.choice(WEB_VARIANTS[rec.channel]),
                int(part),
                round(float(np.clip(bounce, 0.05, 0.95)), 3),
            ))

web = pd.DataFrame(session_rows, columns=["session_date", "source", "sessions", "bounce_rate"])
web = web.sample(frac=1.0, random_state=11).reset_index(drop=True)

# bad extract: a handful of sign-flipped session counts
neg_candidates = web.index[web["sessions"] >= 3].to_numpy()
neg_idx = rng.choice(neg_candidates, size=N_NEGATIVE_SESSION_ROWS, replace=False)
web.loc[neg_idx, "sessions"] = -web.loc[neg_idx, "sessions"]

web = web.sort_values("session_date", kind="stable").reset_index(drop=True)
web.to_csv("web_sessions.csv", index=False)


# 4. channel_spend.xlsx
# two Affiliate months are omitted from the workbook
gap_positions = sorted(rng.choice(range(3, len(MONTHS) - 2), size=N_AFFILIATE_GAPS, replace=False))
affiliate_gaps = {str(MONTHS[i]) for i in gap_positions}

spend_records = []
for rec in truth.itertuples():
    if rec.channel == "Affiliate" and str(rec.month) in affiliate_gaps:
        continue
    as_text = rng.random() < SPEND_TEXT_SHARE
    # round first, then record the rounded figure as the truth, so totals reconcile exactly
    value = round(rec.model_spend) if as_text else round(rec.model_spend, 2)
    spend_records.append(dict(
        month=str(rec.month),
        true_channel=rec.channel,
        month_label=rec.month.to_timestamp().strftime("%b-%y"),
        channel_label=rng.choice(SPEND_VARIANTS[rec.channel]),
        cell=f"${value:,.0f}" if as_text else float(value),
        true_spend=float(value),
    ))

spend_df = pd.DataFrame(spend_records)

wb = Workbook()
ws = wb.active
ws.title = "Spend"
ws.append(["Month", "Channel", "Spend_USD"])
for c in ws[1]:
    c.font = Font(bold=True)
for r in spend_df.itertuples():
    ws.append([r.month_label, r.channel_label, r.cell])
ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 14
wb.save("channel_spend.xlsx")


# 5. Ground-truth summary
true_spend = spend_df.groupby(["month", "true_channel"], as_index=False)["true_spend"].sum()
true_sessions = truth[["month", "channel", "sessions", "orders"]].copy()
true_sessions["month"] = true_sessions["month"].astype(str)
true_sessions = true_sessions.rename(columns={"channel": "true_channel"})

gt = (true_revenue
      .merge(true_spend, on=["month", "true_channel"], how="outer")
      .merge(true_sessions, on=["month", "true_channel"], how="outer")
      .fillna({"revenue": 0.0, "true_spend": 0.0, "sessions": 0, "orders": 0}))

pd.set_option("display.width", 120)
pd.set_option("display.float_format", lambda v: f"{v:,.2f}")

banner("FILES WRITTEN")
print(f"crm_orders.csv       {len(crm_out):>6,} rows  "
      f"({len(orders):,} unique orders + {N_DUPLICATE_IDS} duplicated order_ids)")
print(f"web_sessions.csv     {len(web):>6,} rows  "
      f"({N_NEGATIVE_SESSION_ROWS} rows with negative sessions)")
print(f"channel_spend.xlsx   {len(spend_df):>6,} rows  "
      f"(Affiliate missing: {', '.join(sorted(affiliate_gaps))})")
print(f"nulls in region:     {crm_out['region'].isna().sum():>6,} "
      f"({crm_out['region'].isna().mean():.1%})")
print(f"spend cells as text: {sum(isinstance(v, str) for v in spend_df['cell']):>6,}")

print()
banner("GROUND TRUTH BY CHANNEL (Jan-2025 through Jun-2026)")
by_channel = gt.groupby("true_channel").agg(
    spend=("true_spend", "sum"),
    revenue=("revenue", "sum"),
    orders=("orders", "sum"),
    sessions=("sessions", "sum"),
).reindex(CHANNELS.keys())
by_channel["roas"] = by_channel["revenue"] / by_channel["spend"]
by_channel["aov"] = by_channel["revenue"] / by_channel["orders"]
by_channel["cvr"] = by_channel["orders"] / by_channel["sessions"]
print(by_channel.to_string(formatters={
    "spend": "{:,.0f}".format, "revenue": "{:,.0f}".format,
    "orders": "{:,.0f}".format, "sessions": "{:,.0f}".format,
    "roas": "{:.2f}x".format, "aov": "${:,.0f}".format, "cvr": "{:.2%}".format,
}))
tot = by_channel.sum()
print(f"{'TOTAL':<12}{tot['spend']:>12,.0f}{tot['revenue']:>14,.0f}"
      f"{tot['orders']:>10,.0f}{tot['sessions']:>12,.0f}"
      f"{tot['revenue'] / tot['spend']:>10.2f}x")
print("NOTE: Affiliate spend excludes the two omitted months, so its ROAS here "
      "reads high on purpose.")

print()
banner("GROUND TRUTH BY MONTH")
by_month = gt.groupby("month").agg(
    spend=("true_spend", "sum"),
    revenue=("revenue", "sum"),
    orders=("orders", "sum"),
    sessions=("sessions", "sum"),
)
by_month["roas"] = by_month["revenue"] / by_month["spend"]
print(by_month.to_string(formatters={
    "spend": "{:,.0f}".format, "revenue": "{:,.0f}".format,
    "orders": "{:,.0f}".format, "sessions": "{:,.0f}".format,
    "roas": "{:.2f}x".format,
}))

print()
banner("TRUE MONTHLY REVENUE BY CHANNEL")
print(gt.pivot_table(index="month", columns="true_channel", values="revenue", aggfunc="sum")
        .reindex(columns=list(CHANNELS.keys()))
        .to_string(float_format=lambda v: f"{v:,.0f}"))

print()
banner("TRUE MONTHLY SPEND BY CHANNEL (blank = omitted from the workbook)")
print(true_spend.pivot_table(index="month", columns="true_channel", values="true_spend", aggfunc="sum")
        .reindex(columns=list(CHANNELS.keys()))
        .to_string(float_format=lambda v: f"{v:,.0f}", na_rep="--"))

lost_sessions = int(web.loc[web["sessions"] < 0, "sessions"].abs().sum())
dupe_delta = float(dupes["revenue"].sum() - orders.loc[dupe_idx, "revenue"].sum())

print()
banner("CLEANING NOTES")
print("To reproduce the numbers above: map the name variants to the five true "
      "channels, drop duplicate order_ids, drop the negative session rows, and "
      "strip '$' and ',' from Spend_USD before casting to float.")
print("- Spend reconciles exactly.")
print(f"- Revenue lands within ~{abs(dupe_delta):,.0f} USD "
      f"({abs(dupe_delta) / by_channel['revenue'].sum():.2%}) of truth depending on "
      f"which copy of each duplicated order_id you keep.")
print(f"- Dropping the negative session rows loses {lost_sessions:,} sessions "
      f"({lost_sessions / by_channel['sessions'].sum():.2%}); taking abs() instead "
      f"recovers the true totals exactly.")