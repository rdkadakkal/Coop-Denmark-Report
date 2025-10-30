# app.py
import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime, date, time, timedelta
from zoneinfo import ZoneInfo

st.set_page_config(page_title="Terminal Data Quality Report", layout="wide")

st.title("📦 Terminal Data Quality Report (CET)")
st.caption("Upload your CSV, pick a date range, and download a formatted Excel report.")

# ----------------------------
# UI Inputs
# ----------------------------
uploaded_csv = st.file_uploader("Upload the main CSV", type=["csv"])

# Default origins per your spec (editable)
default_origins = ["Brondby", "Hasselager", "Lineage", "Odense", "Hilton"]
origins = st.multiselect(
    "Origins to include",
    options=default_origins,
    default=default_origins,
    help="You can add/remove origins if needed. Mojibake like 'Br√∏ndby DC' is normalized to 'Brondby'."
)

# Date range (CET)
col_a, col_b = st.columns(2)
with col_a:
    start_date = st.date_input("Start Date (CET)", value=date(2024, 9, 29))
with col_b:
    end_date = st.date_input("End Date (CET)", value=date(2024, 10, 6))

if end_date < start_date:
    st.error("End Date must be on or after Start Date.")
    st.stop()

# Process button
run = st.button("Generate Excel Report", type="primary", use_container_width=True)

# ----------------------------
# Helpers
# ----------------------------
NEEDED_HEADERS_HIGHLIGHT = {
    "Stop type",
    "Stop actual arrival time",
    "Stop arrival delta (minutes)",
    "Shipment ID",
    "Origin",
    "Destination initial planned arrival time",
}

TDQ_COL_ORDER = [
    "Origin",
    "Registration Rate",      # moved before Planned Deliveries
    "Delivery Precision",     # moved before Planned Deliveries
    "Planned Deliveries",
    "Actual Deliveries",
    "Delayed Deliveries",
]

def insert_column_after(df, after_col, new_col, values):
    cols = list(df.columns)
    if after_col not in cols:
        df[new_col] = values
        return df
    i = cols.index(after_col) + 1
    df.insert(i, new_col, values)
    return df

def parse_datetime_utc(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, utc=True, errors="coerce")

def normalize_origin(x: str) -> str:
    if not isinstance(x, str):
        return ""
    s = x.strip()
    s_lower = s.lower()
    # Fix mojibake & accents -> brondby
    s_lower = s_lower.replace("br√∏ndby", "brondby").replace("brøndby", "brondby")
    # Drop common suffix like " dc"
    s_lower = re.sub(r"\s+dc\b", "", s_lower)
    if "brondby" in s_lower:
        return "Brondby"
    if "hasselager" in s_lower:
        return "Hasselager"
    if "lineage" in s_lower:
        return "Lineage"
    if "odense" in s_lower:
        return "Odense"
    if "hilton" in s_lower:
        return "Hilton"
    return s  # keep original otherwise

def is_nonblank(val):
    """Treat '', ' ', np.nan, 'NaT', and whitespace-only as blank."""
    if val is None:
        return False
    s = str(val)
    if s.lower() in ("nan", "nat"):
        return False
    return s.strip() != ""

def filter_by_cet_range(cet_series, start_d: date, end_d: date):
    """Keep rows whose CET timestamp is within [start_d 00:00:00, end_d 23:59:59.999999]."""
    if cet_series.isna().all():
        return cet_series.notna() & False  # empty mask if no dates
    start_dt = datetime.combine(start_d, time.min).replace(tzinfo=ZoneInfo("Europe/Copenhagen"))
    # End inclusive: push to the end of the day
    end_dt = datetime.combine(end_d, time.max).replace(tzinfo=ZoneInfo("Europe/Copenhagen"))
    return (cet_series >= pd.Timestamp(start_dt)) & (cet_series <= pd.Timestamp(end_dt))

# ----------------------------
# Main action
# ----------------------------
if run:
    if uploaded_csv is None:
        st.error("Please upload the main CSV first.")
        st.stop()

    # Load CSV
    df = pd.read_csv(uploaded_csv)

    # Ensure required columns exist (create blank if missing to avoid hard errors)
    for col in NEEDED_HEADERS_HIGHLIGHT:
        if col not in df.columns:
            df[col] = pd.Series([np.nan] * len(df))

    # Build CET column next to Destination initial planned arrival time
    utc_series = df["Destination initial planned arrival time"].apply(parse_datetime_utc)
    cet_series = utc_series.dt.tz_convert(ZoneInfo("Europe/Copenhagen"))
    df = insert_column_after(df, "Destination initial planned arrival time", "CET", cet_series)

    # Filter rows by CET ∈ [start_date, end_date] inclusive
    mask = filter_by_cet_range(df["CET"], start_date, end_date)
    df_filtered = df.loc[mask].copy()

    # Normalize origins (mojibake fixes etc.)
    df_filtered["_OriginNorm"] = df_filtered["Origin"].apply(normalize_origin)

    # Compute Terminal Data Quality metrics
    results = []
    for origin in origins:
        sub = df_filtered[df_filtered["_OriginNorm"] == origin]

        planned = len(sub)
        # Actual deliveries: count non-blank "Stop actual arrival time"
        actual_times = sub["Stop actual arrival time"]
        actual = int(actual_times.apply(is_nonblank).sum())

        # Delayed deliveries: positive values in Stop arrival delta (minutes)
        delta = pd.to_numeric(sub["Stop arrival delta (minutes)"], errors="coerce")
        delayed = int((delta > 0).sum())

        reg_rate = (actual / planned) if planned > 0 else np.nan
        del_prec = (1 - (delayed / actual)) if actual > 0 else np.nan

        results.append({
            "Origin": origin,
            "Registration Rate": reg_rate,
            "Delivery Precision": del_prec,
            "Planned Deliveries": planned,
            "Actual Deliveries": actual,
            "Delayed Deliveries": delayed,
        })

    tdq_df = pd.DataFrame(results, columns=TDQ_COL_ORDER)

    # Header table ("Terminal Quality") - only once
    header_table = pd.DataFrame({
        "Terminal Quality": ["Start Date", "End Date", "Carrier"],
        "Value": [start_date.strftime("%d-%b"), end_date.strftime("%d-%b"), "All"]
    })

    # Prepare Excel in memory
    # Make timezone-naive for Excel compatibility
    df_out = df_filtered.copy()
    for c in df_out.columns:
        if pd.api.types.is_datetime64tz_dtype(df_out[c]):
            df_out[c] = df_out[c].dt.tz_localize(None)

    output = io.BytesIO()
    main_sheet_name = "Main (Filtered to CET Range)"
    tdq_sheet = "Terminal Data Quality"

    with pd.ExcelWriter(output, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd hh:mm:ss",
                        date_format="yyyy-mm-dd") as writer:
        # Main data
        df_out.to_excel(writer, index=False, sheet_name=main_sheet_name)
        wb = writer.book
        ws_main = writer.sheets[main_sheet_name]

        # Header formats
        header_format_default = wb.add_format({
            "bold": True, "text_wrap": True, "valign": "top", "border": 1
        })
        header_format_yellow = wb.add_format({
            "bold": True, "text_wrap": True, "valign": "top", "border": 1, "bg_color": "#FFF59D"
        })

        # Re-write header row to apply highlight formatting
        for col_idx, col_name in enumerate(df_out.columns):
            fmt = header_format_yellow if col_name in NEEDED_HEADERS_HIGHLIGHT else header_format_default
            ws_main.write(0, col_idx, col_name, fmt)

        # Column widths
        for idx, col in enumerate(df_out.columns):
            width = max(12, min(40, int(df_out[col].astype(str).str.len().quantile(0.95))))
            ws_main.set_column(idx, idx, width)

        # Terminal Data Quality sheet
        header_table.to_excel(writer, index=False, sheet_name=tdq_sheet, startrow=0)
        tdq_df.to_excel(writer, index=False, sheet_name=tdq_sheet, startrow=len(header_table) + 2)
        ws_tdq = writer.sheets[tdq_sheet]

        # Style header table
        bold_fmt = wb.add_format({"bold": True, "bg_color": "#E0E0E0", "border": 1})
        val_fmt = wb.add_format({"border": 1})
        for i in range(len(header_table)):
            ws_tdq.write(i, 0, header_table.iloc[i, 0], bold_fmt)
            ws_tdq.write(i, 1, header_table.iloc[i, 1], val_fmt)

        # Style TDQ headers and percentages
        for col_idx, col_name in enumerate(tdq_df.columns):
            ws_tdq.write(len(header_table) + 2, col_idx, col_name, header_format_default)

        pct_fmt = wb.add_format({"num_format": "0.00%"})
        # Set % formatting on the correct columns
        reg_idx = tdq_df.columns.get_loc("Registration Rate")
        prec_idx = tdq_df.columns.get_loc("Delivery Precision")
        ws_tdq.set_column(reg_idx, reg_idx, 18, pct_fmt)
        ws_tdq.set_column(prec_idx, prec_idx, 18, pct_fmt)

        # Widths for others
        for name, i in zip(tdq_df.columns, range(len(tdq_df.columns))):
            if name not in ["Registration Rate", "Delivery Precision"]:
                ws_tdq.set_column(i, i, 20)

    output.seek(0)

    st.subheader("📄 Terminal Data Quality (Preview)")
    st.dataframe(tdq_df, use_container_width=True)

    st.download_button(
        label="⬇️ Download Excel Report",
        data=output,
        file_name=f"Terminal_Data_Quality_{start_date:%d%b}-{end_date:%d%b}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("""
**Notes**
- Times are converted from UTC to CET/CEST using `Europe/Copenhagen`.
- The CET filter applies to the range you select (inclusive).
- “Actual Deliveries” counts only non-blank values in **Stop actual arrival time**.
- “Carrier” is shown **once** in the header summary.  
""")
