"""Streamlit web app: upload messy order exports and download one clean workbook.

Run locally:  streamlit run app.py
"""
import tempfile
from pathlib import Path

import streamlit as st

from cleaner.output import report_table, workbook_bytes
from cleaner.pipeline import clean_orders
from cleaner.reader import RAW_DIR, read_all
from cleaner.values import SOURCE_DATE_FORMATS

# What the user sees in the date format dropdown -> the format the cleaner uses.
DATE_FORMAT_CHOICES = {
    "Real Excel dates": None,
    "Day first, dots (24.03.2026)": "%d.%m.%Y",
    "Month first, slashes (3/24/2026)": "%m/%d/%Y",
    "Day first, slashes (24/03/2026)": "%d/%m/%Y",
    "ISO (2026-03-24)": "%Y-%m-%d",
}
FORMAT_LABELS = {fmt: label for label, fmt in DATE_FORMAT_CHOICES.items()}

SAMPLE = "Try it with the sample files"
UPLOAD = "Upload your own files"


def sample_files():
    """The three demo exports as (file name, bytes) pairs."""
    return [(p.name, p.read_bytes()) for p in sorted(RAW_DIR.glob("*.xlsx"))]


def run_cleaner(files, date_formats):
    """Writes the files to a temporary folder and runs the normal pipeline on it.
    Returns (raw, orders, rejected, report)."""
    with tempfile.TemporaryDirectory() as folder:
        for name, data in files:
            Path(folder, name).write_bytes(data)
        raw = read_all(folder)
    orders, rejected, report = clean_orders(raw, date_formats)
    return raw, orders, rejected, report


def as_text(df):
    """Mixed text/number columns (like the raw data) shown as plain text."""
    return df.astype(object).where(df.notna(), "").astype(str)


st.set_page_config(page_title="Excel Data Cleaner", page_icon="🧹", layout="wide")

st.title("Excel Data Cleaner")
st.write(
    "Merges messy Excel order exports into one clean workbook. "
    "Nothing is deleted silently: every rejected row is listed with its reason "
    "and its row number in the original file."
)

source = st.radio("Data", [SAMPLE, UPLOAD], horizontal=True, label_visibility="collapsed")

if source == SAMPLE:
    files = sample_files()
    st.caption(
        "Three exports of a fictional retailer: two stores and an online shop, "
        "each with its own column names, date format and mistakes."
    )
    columns = st.columns(len(files))
    for column, (name, data) in zip(columns, files):
        column.download_button(f"⬇ {name}", data, file_name=name, use_container_width=True)
else:
    uploaded = st.file_uploader("Excel exports (.xlsx)", type="xlsx", accept_multiple_files=True)
    files = [(f.name, f.getvalue()) for f in uploaded]
    st.caption(
        "Columns are matched by name (e.g. 'Order ID', 'order_id' or 'ID'). "
        "Products and cities are checked against the demo catalogue."
    )

if not files:
    st.info("Upload one or more .xlsx files to start.")
    st.stop()

st.subheader("How are dates written in each file?")
st.caption("A date like 2/5/2026 means a different day in the US and in Europe, so the cleaner asks instead of guessing.")
date_formats = {}
for column, (name, _) in zip(st.columns(len(files)), files):
    default = FORMAT_LABELS.get(SOURCE_DATE_FORMATS.get(name), "Real Excel dates")
    label = column.selectbox(name, list(DATE_FORMAT_CHOICES),
                             index=list(DATE_FORMAT_CHOICES).index(default))
    date_formats[name] = DATE_FORMAT_CHOICES[label]

if not st.button("Clean data", type="primary"):
    st.stop()

try:
    raw, orders, rejected, report = run_cleaner(files, date_formats)
except (ValueError, FileNotFoundError) as error:
    st.error(f"The files could not be processed: {error}")
    st.stop()

metrics = st.columns(4)
metrics[0].metric("Rows read", report["rows_read"])
metrics[1].metric("Clean orders", report["rows_clean"])
metrics[2].metric("Rejected rows", report["rows_rejected"])
metrics[3].metric("Duplicates removed", report["duplicates_removed"])

st.download_button(
    "⬇ Download clean_orders.xlsx",
    workbook_bytes(orders, rejected, report),
    file_name="clean_orders.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

orders_tab, rejected_tab, report_tab, raw_tab = st.tabs(
    ["Orders", "Rejected", "Report", "Raw data (before)"])
with orders_tab:
    st.dataframe(orders, use_container_width=True, hide_index=True)
with rejected_tab:
    st.dataframe(as_text(rejected), use_container_width=True, hide_index=True)
with report_tab:
    st.dataframe(report_table(report, orders), use_container_width=True, hide_index=True)
with raw_tab:
    st.caption("The rows exactly as they were read from the files, before any cleaning.")
    st.dataframe(as_text(raw), use_container_width=True, hide_index=True)
