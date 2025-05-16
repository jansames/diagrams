"""streamlit_app.py

Classify electricity‑consumption Excel sheets as:
  • 1‑D hourly (8 760 / 8 784 rows)
  • 1‑D 15‑minute (35 040 / 35 136 rows)
  • 2‑D hourly (365 / 366 rows × 24 cols)
  • 2‑D 15‑minute (365 / 366 rows × 96 cols)

On failure, a *very* specific error is returned telling the user exactly what
looks wrong (row count, column count, date‑parsing issues, etc.).
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

# ──────────────────────────────── Constants ────────────────────────────────── #
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

# ───────────────────────────── Helper utilities ───────────────────────────── #

def _safe_to_datetime(series: pd.Series) -> tuple[pd.Series | None, str | None]:
    """Return (datetime_series, error_msg).  *error_msg* is *None* on success."""

    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().any():
        return None, "Failed to parse dates/timestamps in the first column."
    return dt, None


def _explain_mismatch(found: int, expected: list[int]) -> str:
    exp_str = ", ".join(str(e) for e in expected)
    return f"Found **{found}**, but expected one of **{exp_str}**."


# ───────────────────────────── Core detection ─────────────────────────────── #

def detect_format(df: pd.DataFrame) -> tuple[str, str]:
    """Return (label, detail).  *label* begins with "Error" on failure."""

    # Strip empty rows/cols early for robust counting
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    rows, cols = df.shape

    # ───────────── 1‑D candidate ───────────── #
    if rows in ROWS_1D and cols >= 2:
        ts, err = _safe_to_datetime(df.iloc[:, 0])
        if ts is None:
            return "Error - timestamp parsing", err  # type: ignore[return-value]
        if ts.dt.hour.max() > 0 or ts.dt.minute.max() > 0:
            if rows in {HOUR_ROWS_NONLEAP, HOUR_ROWS_LEAP}:
                return "1D hourly", "Row count matches a full year of hourly data."
            if rows in {Q15_ROWS_NONLEAP, Q15_ROWS_LEAP}:
                return "1D 15 minutes", "Row count matches a full year of 15‑minute data."
            msg = _explain_mismatch(rows, [*ROWS_1D])
            return "Error - unexpected row count for 1D", msg  # type: ignore[return-value]

    # ───────────── 2‑D candidate ───────────── #
    if rows in ROWS_2D and cols >= 2:
        dates, err = _safe_to_datetime(df.iloc[:, 0])
        if dates is None:
            return "Error - date parsing", err  # type: ignore[return-value]
        if not (dates.dt.hour.eq(0).all() and dates.dt.minute.eq(0).all()):
            return "Error - first column not pure dates", "First column includes non‑midnight timestamps."

        n_interval_cols = cols - 1
        if n_interval_cols == 24:
            return "2D hourly", "365/366 rows × 24 interval columns detected."
        if n_interval_cols == 96:
            return "2D 15 minutes", "365/366 rows × 96 interval columns detected."
        msg = _explain_mismatch(n_interval_cols, [24, 96])
        return "Error - unexpected number of interval columns", msg  # type: ignore[return-value]

    # ───────────── Unknown ───────────── #
    detail = (
        "Could not match known patterns.\n\n"
        f"• Row count after cleaning: **{rows}**\n"
        f"• Total columns: **{cols}**"
    )
    return "Error - unrecognized format", detail  # type: ignore[return-value]


# ───────────────────────────── Streamlit UI ───────────────────────────────── #

def main() -> None:  # noqa: D401
    st.set_page_config(page_title="Electricity Diagram Format Recognizer", page_icon="⚡")
    st.title("⚡ Electricity Diagram Format Recognizer")

    # Check openpyxl availability
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        st.error(
            "**openpyxl** is required to read .xlsx files. Install it with `pip install openpyxl`"
            " (or add it to *requirements.txt*) and restart."
        )
        st.stop()

    uploaded = st.file_uploader("Upload an XLSX workbook", type=["xlsx", "xls"], key="uploader")
    if uploaded is None:
        st.info("⬆️ Drag & drop or browse to upload an Excel file.")
        st.stop()

    # Read workbook
    try:
        xls = pd.ExcelFile(uploaded)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Could not open file: {exc}")
        st.stop()

    # Select sheet (default: first)
    sheet = xls.sheet_names[0]
    if len(xls.sheet_names) > 1:
        sheet = st.selectbox("Select sheet", xls.sheet_names, key="sheet_select")

    try:
        df = xls.parse(sheet)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Failed to parse sheet: {exc}")
        st.stop()

    st.subheader("Data preview (first 5 rows)")
    st.dataframe(df.head())

    label, info = detect_format(df)
    if label.startswith("Error"):
        st.error(label)
        st.write(info)
    else:
        st.success(f"Detected format: **{label}**")
        st.write(info)


if __name__ == "__main__":
    main()
