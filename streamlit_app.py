"""streamlit_app.py

Classify electricity‑consumption Excel sheets even when they contain several
years of data.  The app focuses on the first **complete** calendar year it can
find and labels the sheet as one of:

* 1‑D hourly
* 1‑D 15 minutes
* 2‑D hourly
* 2‑D 15 minutes

— with detailed, user‑friendly error explanations if something doesn’t look
right (row‑ or column‑count mismatch, parsing problems, etc.).
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

# ──────────────────────────────── Constants ────────────────────────────────── #
HOUR_ROWS_NONLEAP = 8_760
HOUR_ROWS_LEAP = 8_784
Q15_ROWS_NONLEAP = 35_040
Q15_ROWS_LEAP = 35_136

FULL_YEAR_ROWS_1D = {
    HOUR_ROWS_NONLEAP,
    HOUR_ROWS_LEAP,
    Q15_ROWS_NONLEAP,
    Q15_ROWS_LEAP,
}

# ───────────────────────────── Helper utilities ───────────────────────────── #

def _safe_to_datetime(series: pd.Series) -> tuple[pd.Series | None, str | None]:
    """Convert to datetime; return (*dt*, *err_msg*)."""

    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().any():
        return None, "Failed to parse valid dates/timestamps in the first column."
    return dt, None


def _explain_mismatch(found: int, expected: list[int | str]) -> str:
    exp_str = ", ".join(str(e) for e in expected)
    return f"Found **{found}**, expected one of **{exp_str}**."


# ───────────────────────────── Core detection ─────────────────────────────── #

def detect_format(df: pd.DataFrame) -> tuple[str, str]:
    """Return (*label*, *detail*).  *label* starts with "Error" on failure."""

    # Remove completely empty rows/cols (headers are not part of *df.shape*)
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    total_rows, total_cols = df.shape

    # First‑column must be datetime‑convertible for *all* supported layouts
    dt, err = _safe_to_datetime(df.iloc[:, 0])
    if dt is None:
        return "Error - date parsing", err  # type: ignore[return-value]

    # Detect 1‑D vs 2‑D based on presence of time information in timestamps
    has_time_detail = (dt.dt.hour.max() > 0) or (dt.dt.minute.max() > 0)

    # ────────────────────────── 1‑D candidate ────────────────────────────── #
    if has_time_detail:
        # Group by calendar year and look for a *complete* one
        year_counts = dt.dt.year.value_counts().sort_index()
        for year, count in year_counts.items():
            if count in FULL_YEAR_ROWS_1D:
                if count in {HOUR_ROWS_NONLEAP, HOUR_ROWS_LEAP}:
                    return (
                        "1D hourly",
                        f"Detected {count} rows for {year} – full hourly dataset.",
                    )
                if count in {Q15_ROWS_NONLEAP, Q15_ROWS_LEAP}:
                    return (
                        "1D 15 minutes",
                        f"Detected {count} rows for {year} – full 15‑minute dataset.",
                    )
        # If we get here: no complete year found
        detail = (
            "1‑D style timestamps, but no calendar year has a full set of rows.\n\n"
            + "Row counts per year: "
            + ", ".join(f"{y}: {c}" for y, c in year_counts.items())
        )
        return "Error - incomplete 1D year", detail  # type: ignore[return-value]

    # ────────────────────────── 2‑D candidate ────────────────────────────── #
    # Must be midnight‑only dates
    if not (dt.dt.hour.eq(0).all() and dt.dt.minute.eq(0).all()):
        return (
            "Error - first column contains times",
            "2‑D diagrams should have *dates only* (00:00) in first column.",
        )  # type: ignore[return-value]

    # Numeric interval columns (ignore text columns like *Remark*, etc.)
    numeric_cols = [c for c in df.columns[1:] if pd.api.types.is_numeric_dtype(df[c])]
    n_numeric = len(numeric_cols)

    # Determine granularity based on number of numeric columns (≥24 → hourly, ≥96 → 15‑min)
    if n_numeric < 24:
        return (
            "Error - not enough interval columns",
            _explain_mismatch(n_numeric, ["≥24 for hourly", "≥96 for 15‑minute"]),
        )  # type: ignore[return-value]

    granularity = "hourly" if n_numeric < 96 else "15 minutes"
    expected_cols = 24 if granularity == "hourly" else 96

    # Pick the first calendar year that has ≥365 rows
    year_counts = dt.dt.year.value_counts().sort_index()
    for year, count in year_counts.items():
        if count >= 365:  # leap‑year OK (366)
            detail_rows = f"Using {count} rows for year {year}."
            break
    else:  # no break
        return (
            "Error - no full year of rows",
            "The sheet has dates, but none of the years contains ≥365 rows.",
        )  # type: ignore[return-value]

    # Additional detail about extra numeric columns (totals, etc.)
    extra_cols = n_numeric - expected_cols
    extra_note = (
        " (" + ("+" if extra_cols > 0 else "") + f"{extra_cols} extra column(s) ignored)"
        if extra_cols != 0
        else ""
    )

    return (
        f"2D {granularity}",
        f"{detail_rows} Detected {n_numeric} numeric columns – first {expected_cols} treated as "
        f"interval data{extra_note}.",
    )


# ───────────────────────────── Streamlit UI ───────────────────────────────── #

def main() -> None:  # noqa: D401
    st.set_page_config(page_title="Electricity Diagram Format Recognizer", page_icon="⚡")
    st.title("⚡ Electricity Diagram Format Recognizer")

    # Verify openpyxl availability once for friendlier UX
    try:
        import openpyxl  # noqa: F401 – presence check only
    except ModuleNotFoundError:
        st.error(
            "**openpyxl** is required for .xlsx files. Install with `pip install openpyxl`"
            " (or add it to *requirements.txt*) and restart."
        )
        st.stop()

    uploaded = st.file_uploader("Upload an XLSX workbook", type=["xlsx", "xls"], key="uploader")
    if uploaded is None:
        st.info("⬆️ Drag & drop or browse to upload an Excel file.")
        st.stop()

    # Attempt to read workbook
    try:
        xls = pd.ExcelFile(uploaded)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Could not open file: {exc}")
        st.stop()

    # Sheet picker
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

    label, detail = detect_format(df)
    if label.startswith("Error"):
        st.error(label)
        st.write(detail)
    else:
        st.success(f"Detected format: **{label}**")
        st.write(detail)


if __name__ == "__main__":
    main()
