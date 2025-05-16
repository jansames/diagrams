"""streamlit_app.py

An interactive Streamlit utility that classifies electricity‑consumption Excel
sheets as either 1‑D or 2‑D diagrams and determines their granularity (hourly
or 15‑minute) over a complete calendar year.

Returns one of:
    - "1D hourly"
    - "1D 15 minutes"
    - "2D hourly"
    - "2D 15 minutes"
    - "Error - …" (with explanation)
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

# ------------------------------ Constants ---------------------------------- #
HOUR_ROWS_NONLEAP = 8_760
HOUR_ROWS_LEAP = 8_784
Q15_ROWS_NONLEAP = 35_040
Q15_ROWS_LEAP = 35_136

ROWS_1D = {
    HOUR_ROWS_NONLEAP,
    HOUR_ROWS_LEAP,
    Q15_ROWS_NONLEAP,
    Q15_ROWS_LEAP,
}
ROWS_2D = {365, 366}


# ------------------------------ Helpers ------------------------------------ #

def _safe_to_datetime(series: pd.Series) -> pd.Series | None:
    """Convert *series* to datetime; return *None* if any entry fails."""

    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().any():
        return None
    return dt


# ------------------------------ Core logic --------------------------------- #

def detect_format(df: pd.DataFrame) -> str:
    """Detect diagram layout & granularity.

    Parameters
    ----------
    df : pandas.DataFrame
        Raw DataFrame from Excel.

    Returns
    -------
    str
        One of the labels listed in the module docstring.
    """

    # Strip completely empty rows/cols for robust detection
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    # --------------------------- 1‑D candidate ----------------------------- #
    if df.shape[0] in ROWS_1D and df.shape[1] >= 2:
        ts = _safe_to_datetime(df.iloc[:, 0])
        if ts is not None and (ts.dt.hour.max() > 0 or ts.dt.minute.max() > 0):
            rows = len(df)
            if rows in {HOUR_ROWS_NONLEAP, HOUR_ROWS_LEAP}:
                return "1D hourly"
            if rows in {Q15_ROWS_NONLEAP, Q15_ROWS_LEAP}:
                return "1D 15 minutes"
            return (
                "Error - 1D detected but unexpected row count; expected 8760/8784 or"
                " 35040/35136."
            )

    # --------------------------- 2‑D candidate ----------------------------- #
    if df.shape[0] in ROWS_2D and df.shape[1] >= 2:
        dates = _safe_to_datetime(df.iloc[:, 0])
        if dates is not None and dates.dt.hour.eq(0).all() and dates.dt.minute.eq(0).all():
            numeric_cols = [c for c in df.columns[1:] if pd.api.types.is_numeric_dtype(df[c])]
            n = len(numeric_cols)
            if n == 24:
                return "2D hourly"
            if n == 96:
                return "2D 15 minutes"
            return (
                f"Error - 2D layout detected but found {n} numeric columns; expected 24 "
                "(hourly) or 96 (15‑min)."
            )

    # --------------------------- Failure ----------------------------------- #
    return "Error - unrecognized format"


# ----------------------------- Streamlit UI -------------------------------- #

def main() -> None:  # noqa: D401
    st.set_page_config(page_title="Electricity Diagram Format Recognizer", page_icon="⚡")
    st.title("⚡ Electricity Diagram Format Recognizer")

    # Ensure openpyxl is available before proceeding
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        st.error(
            "**openpyxl** is required to read .xlsx files. Install it with `pip install "
            "openpyxl` (or add it to *requirements.txt*) and restart."
        )
        st.stop()

    uploaded = st.file_uploader("Upload an XLSX workbook", type=["xlsx", "xls"], key="uploader")
    if uploaded is None:
        st.info("⬆️ Drag & drop or browse to upload an Excel file.")
        st.stop()

    # Attempt to read the workbook
    try:
        xls = pd.ExcelFile(uploaded)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Could not open file: {exc}")
        st.stop()

    # Sheet selection UI
    sheet = xls.sheet_names[0]
    if len(xls.sheet_names) > 1:
        sheet = st.selectbox("Select sheet", xls.sheet_names, key="sheet_select")

    # Parse selected sheet
    try:
        df = xls.parse(sheet)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Failed to parse sheet: {exc}")
        st.stop()

    # Data preview
    st.subheader("Data preview (first 5 rows)")
    st.dataframe(df.head())

    # Format detection
    result = detect_format(df)
    if result.startswith("Error"):
        st.error(result)
    else:
        st.success(f"Detected format: **{result}**")


if __name__ == "__main__":
    main()
