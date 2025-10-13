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
    "Destination initial planned
