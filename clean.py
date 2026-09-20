#!/usr/bin/env python3
"""Clean the three raw marketing files and merge them into one month x channel table for Tableau."""

import os

import pandas as pd

RAW_DIR = "original_data"
CLEAN_DIR = "cleaned_data"
os.makedirs(CLEAN_DIR, exist_ok=True)

CANONICAL_CHANNELS = ["Google Ads", "Meta", "Email", "Organic", "Affiliate"]


def normalize_channel(raw: str) -> str:
    """Map any of the three files' naming conventions to one canonical channel."""
    s = str(raw).strip().lower()

    if "organic" in s or "seo" in s:
        return "Organic"
    if "affiliate" in s or "partner" in s or s == "aff":
        return "Affiliate"
    if "meta" in s or "facebook" in s or "fb" in s:
        return "Meta"
    if "email" in s or "e-mail" in s or "klaviyo" in s:
        return "Email"
    if "google" in s or "adwords" in s or "cpc" in s:
        return "Google Ads"

    raise ValueError(f"Unmapped channel value: {raw!r}")


def banner(title):
    print("=" * 70)
    print(title)
    print("=" * 70)


def clean_money(v) -> float:
    """Strip '$' and ',' from text money cells and cast everything to float."""
    if isinstance(v, str):
        return float(v.replace("$", "").replace(",", "").strip())
    return float(v)


# 1. crm_orders.csv
crm = pd.read_csv(os.path.join(RAW_DIR, "crm_orders.csv"))
crm_rows_in = len(crm)

crm["channel"] = crm["channel"].map(normalize_channel)
crm["order_date"] = pd.to_datetime(crm["order_date"], format="%m/%d/%Y")
crm["month"] = crm["order_date"].dt.to_period("M").astype(str)
crm["region"] = crm["region"].fillna("Unknown")

crm_dupes = crm["order_id"].duplicated().sum()
crm = crm.drop_duplicates(subset="order_id", keep="first").reset_index(drop=True)

crm_clean = crm[["order_id", "order_date", "month", "customer_id", "channel", "revenue", "region"]]
crm_clean.to_csv(os.path.join(CLEAN_DIR, "cleaned_crm_orders.csv"), index=False)

# 2. web_sessions.csv
web = pd.read_csv(os.path.join(RAW_DIR, "web_sessions.csv"))
web_rows_in = len(web)

web["channel"] = web["source"].map(normalize_channel)
web["session_date"] = pd.to_datetime(web["session_date"], format="%Y-%m-%d")
web["month"] = web["session_date"].dt.to_period("M").astype(str)

web_negative = int((web["sessions"] < 0).sum())
web["sessions"] = web["sessions"].abs()

web_clean = web[["session_date", "month", "channel", "sessions", "bounce_rate"]]
web_clean.to_csv(os.path.join(CLEAN_DIR, "cleaned_web_sessions.csv"), index=False)

# 3. channel_spend.xlsx
spend = pd.read_excel(os.path.join(RAW_DIR, "channel_spend.xlsx"), sheet_name="Spend")
spend_rows_in = len(spend)

spend["channel"] = spend["Channel"].map(normalize_channel)
spend["spend"] = spend["Spend_USD"].map(clean_money)
spend["month"] = pd.to_datetime(spend["Month"], format="%b-%y").dt.to_period("M").astype(str)

spend_clean = spend[["month", "channel", "spend"]]
spend_clean.to_csv(os.path.join(CLEAN_DIR, "cleaned_channel_spend.csv"), index=False)

# 4. Merge into one month x channel performance table
rev_agg = crm_clean.groupby(["month", "channel"], as_index=False).agg(
    revenue=("revenue", "sum"), orders=("order_id", "count")
)
# several source labels map to one channel, so sum sessions and weight bounce_rate by sessions
web_w = web_clean.assign(bounces=web_clean["sessions"] * web_clean["bounce_rate"])
sess_agg = web_w.groupby(["month", "channel"], as_index=False).agg(
    sessions=("sessions", "sum"), bounces=("bounces", "sum")
)
sess_agg = sess_agg.rename(columns={"bounces": "bounced_sessions"})

spend_agg = spend_clean.groupby(["month", "channel"], as_index=False).agg(spend=("spend", "sum"))

perf = (
    spend_agg.merge(rev_agg, on=["month", "channel"], how="outer")
    .merge(sess_agg, on=["month", "channel"], how="outer")
)
# missing spend stays NaN (unknown), not 0
perf[["revenue", "orders", "sessions", "bounced_sessions"]] = perf[
    ["revenue", "orders", "sessions", "bounced_sessions"]
].fillna(0)

# Tableau-ready: additive measures only (ratios are built in Tableau), a date column, a missing-spend flag
perf["month_start"] = pd.to_datetime(perf["month"] + "-01").dt.strftime("%Y-%m-%d")
perf["spend_missing"] = perf["spend"].isna().astype(int)
perf["bounced_sessions"] = perf["bounced_sessions"].round(2)

perf = perf.sort_values(["month", "channel"]).reset_index(drop=True)
final = perf[[
    "month_start", "month", "channel", "spend", "spend_missing",
    "revenue", "orders", "sessions", "bounced_sessions",
]]
final.to_csv(os.path.join(CLEAN_DIR, "channel_performance.csv"), index=False)

# 4b. Excel copies of every output, with real Excel dates
excel_outputs = {
    "cleaned_crm_orders": crm_clean,
    "cleaned_web_sessions": web_clean,
    "cleaned_channel_spend": spend_clean,
    "channel_performance": final.assign(month_start=pd.to_datetime(final["month_start"])),
}
for name, df in excel_outputs.items():
    with pd.ExcelWriter(
        os.path.join(CLEAN_DIR, f"{name}.xlsx"), engine="openpyxl", datetime_format="yyyy-mm-dd"
    ) as xw:
        df.to_excel(xw, sheet_name=name[:31], index=False)

# 5. Report
banner("CLEANING SUMMARY")
print(f"crm_orders.csv     : {crm_rows_in:,} rows in -> {len(crm_clean):,} out "
      f"({crm_dupes} duplicate order_ids dropped)")
print(f"web_sessions.csv   : {web_rows_in:,} rows in -> {len(web_clean):,} out "
      f"({web_negative} negative session rows sign-corrected, not dropped)")
print(f"channel_spend.xlsx : {spend_rows_in:,} rows in -> {len(spend_clean):,} out "
      f"(90 month x channel slots expected -> {90 - spend_rows_in} missing, still missing after cleaning)")

print()
banner("CHANNEL TOTALS (18 months, Jan-2025 - Jun-2026)")
by_channel = perf.groupby("channel").agg(
    spend=("spend", "sum"), revenue=("revenue", "sum"),
    orders=("orders", "sum"), sessions=("sessions", "sum"),
).reindex(CANONICAL_CHANNELS)
# ROAS only over months with a spend record, or revenue without spend inflates it
rev_with_spend = perf[perf["spend"].notna()].groupby("channel")["revenue"].sum()
by_channel["roas"] = rev_with_spend / by_channel["spend"]
by_channel["aov"] = by_channel["revenue"] / by_channel["orders"]
by_channel["cvr"] = by_channel["orders"] / by_channel["sessions"]
pd.set_option("display.width", 120)
print(by_channel.to_string(formatters={
    "spend": "{:,.0f}".format, "revenue": "{:,.0f}".format,
    "orders": "{:,.0f}".format, "sessions": "{:,.0f}".format,
    "roas": "{:.2f}x".format, "aov": "${:,.0f}".format, "cvr": "{:.2%}".format,
}))

print()
print("Files written (each as .csv and .xlsx): cleaned_crm_orders, cleaned_web_sessions, "
      "cleaned_channel_spend, channel_performance")
