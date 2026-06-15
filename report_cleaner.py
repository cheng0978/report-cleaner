"""
Report Cleaner — automated Excel/CSV report builder
====================================================

A small Streamlit app that takes a messy Excel or CSV file and turns it into a
clean, summarized report in a few clicks:

    1. Upload a .csv / .xlsx file
    2. The app cleans it (trims spaces, drops blank rows, fixes types)
    3. Pick a column to group by and a numeric column to summarize
    4. Get totals, averages, a chart, and a downloadable cleaned report

Run locally with:
    streamlit run report_cleaner.py

Built as a portfolio piece to demonstrate practical Python data automation:
file I/O, pandas data cleaning, grouping/aggregation, and a simple UI.
Author: Zhi Cheng H.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st


# --------------------------------------------------------------------------- #
# Data helpers (kept separate from the UI so the logic is easy to test/reuse)
# --------------------------------------------------------------------------- #
def load_file(uploaded_file) -> pd.DataFrame:
    """Read an uploaded CSV or Excel file into a DataFrame.

    Tries a couple of common CSV encodings before giving up, since real-world
    files (especially from Excel on Windows) are often not UTF-8.
    """
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        for encoding in ("utf-8", "utf-8-sig", "cp950", "latin-1"):
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=encoding)
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError("Could not read this CSV. Please check the file format.")

    if name.endswith((".xlsx", ".xls")):
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)

    raise ValueError("Unsupported file type. Please upload a .csv or .xlsx file.")


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Clean a raw DataFrame and return it along with a log of what changed.

    The log is shown to the user so the cleaning isn't a black box — useful
    when handing a tool to a non-technical client.
    """
    log: list[str] = []
    original_rows = len(df)

    # 1. Strip whitespace from column names
    df.columns = [str(c).strip() for c in df.columns]

    # 2. Drop fully empty rows and columns
    df = df.dropna(how="all").dropna(axis=1, how="all")
    if len(df) < original_rows:
        log.append(f"Removed {original_rows - len(df)} fully empty row(s).")

    # 3. Trim whitespace in text cells
    text_cols = df.select_dtypes(include="object").columns
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()

    # 4. Remove exact duplicate rows
    before = len(df)
    df = df.drop_duplicates()
    if len(df) < before:
        log.append(f"Removed {before - len(df)} duplicate row(s).")

    # 5. Try to convert object columns that are actually numbers
    for col in text_cols:
        converted = pd.to_numeric(df[col].str.replace(",", "", regex=False),
                                  errors="coerce")
        # only adopt the conversion if most values are valid numbers
        if converted.notna().mean() > 0.8:
            df[col] = converted
            log.append(f"Converted column '{col}' to numbers.")

    if not log:
        log.append("Data was already clean — no changes needed.")

    return df.reset_index(drop=True), log


def build_summary(df: pd.DataFrame, group_col: str, value_col: str) -> pd.DataFrame:
    """Group by one column and summarize a numeric column."""
    summary = (
        df.groupby(group_col)[value_col]
        .agg(Total="sum", Average="mean", Count="count")
        .reset_index()
        .sort_values("Total", ascending=False)
    )
    summary["Average"] = summary["Average"].round(2)
    summary["Total"] = summary["Total"].round(2)
    return summary


def to_excel_bytes(cleaned: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    """Write the cleaned data and summary to a two-sheet Excel file in memory."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="Cleaned Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="Report Cleaner", page_icon="📊", layout="wide")

st.title("📊 Report Cleaner")
st.caption("Turn a messy spreadsheet into a clean, summarized report in seconds.")

uploaded = st.file_uploader("Upload an Excel or CSV file", type=["csv", "xlsx", "xls"])

if uploaded is None:
    st.info("👆 Upload a .csv or .xlsx file to get started. "
            "Don't have one handy? Export any spreadsheet of sales, orders, "
            "or expenses and try it out.")
    st.stop()

# --- Load -----------------------------------------------------------------
try:
    raw = load_file(uploaded)
except ValueError as err:
    st.error(str(err))
    st.stop()

if raw.empty:
    st.error("That file appears to be empty.")
    st.stop()

# --- Clean ----------------------------------------------------------------
cleaned, change_log = clean_data(raw)

st.subheader("1 · Cleaned data")
col_a, col_b = st.columns([3, 1])
with col_a:
    st.dataframe(cleaned, use_container_width=True, height=320)
with col_b:
    st.metric("Rows", len(cleaned))
    st.metric("Columns", cleaned.shape[1])
    st.markdown("**What I cleaned:**")
    for entry in change_log:
        st.write("• " + entry)

# --- Summarize ------------------------------------------------------------
st.subheader("2 · Build a summary")

numeric_cols = cleaned.select_dtypes(include="number").columns.tolist()
all_cols = cleaned.columns.tolist()

if not numeric_cols:
    st.warning("No numeric columns found, so a numeric summary isn't possible. "
               "The cleaned data above can still be downloaded below.")
    summary = None
else:
    c1, c2 = st.columns(2)
    with c1:
        group_col = st.selectbox("Group by", all_cols,
                                 help="e.g. Region, Product, Salesperson")
    with c2:
        value_col = st.selectbox("Summarize (numeric)", numeric_cols,
                                 help="e.g. Revenue, Quantity, Amount")

    summary = build_summary(cleaned, group_col, value_col)
    st.dataframe(summary, use_container_width=True)

    # simple bar chart of the totals
    chart_data = summary.set_index(group_col)["Total"]
    st.bar_chart(chart_data)

# --- Download -------------------------------------------------------------
st.subheader("3 · Download your report")

if summary is not None:
    excel_bytes = to_excel_bytes(cleaned, summary)
    stamp = datetime.now().strftime("%Y%m%d")
    st.download_button(
        "⬇️  Download Excel report (cleaned + summary)",
        data=excel_bytes,
        file_name=f"report_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    csv_bytes = cleaned.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️  Download cleaned data (CSV)",
        data=csv_bytes,
        file_name="cleaned_data.csv",
        mime="text/csv",
    )

st.divider()
st.caption("Built with Python, pandas & Streamlit · a portfolio demo by Zhi Cheng H.")
