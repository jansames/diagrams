"""streamlit_app.py

Detects and classifies electricity‑consumption Excel sheets, with robust error
handling so that unexpected problems surface as *friendly* messages rather
than a generic "Axios 400" in the browser.

## Assumptions
* **1‑D hourly** data are labelled with the *end* of each interval, so the first
  timestamp of a year is **01‑Jan 01:00**.
* **1‑D 15‑minute** data start at **01‑Jan 00:15**.
* **2‑D** sheets have one date per row and ≥24 / ≥96 numeric columns.

The detector tolerates a missing *final midnight* row (–1) and ignores extra
numeric columns in 2‑D layouts.
"""

from __future__ import annotations

import calendar
from datetime import datetime
from typing import Final

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

# ─────────────────────────────── CONSTANTS ───────────────────────────────── #
HOUR_ROWS_NONLEAP: Final = 8_760
HOUR_ROWS_LEAP: Final = 8_784
Q15_ROWS_NONLEAP: Final = 35_040
Q15_ROWS_LEAP: Final = 35_136

ACCEPTABLE_HOURLY = {
    HOUR_ROWS_NONLEAP,
    HOUR_ROWS_NONLEAP - 1,
    HOUR_ROWS_LEAP,
    HOUR_ROWS_LEAP - 1,
}
ACCEPTABLE_15M = {
    Q15_ROWS_NONLEAP,
    Q15_ROWS_NONLEAP - 1,
    Q15_ROWS_LEAP,
    Q15_ROWS_LEAP - 1,
}

MAX_UPLOAD_MB: Final = 50  # hard‑stop for very large files (>50 MB)

# ─────────────────────────────── HELPERS ─────────────────────────────────── #

def _safe_to_datetime(series: pd.Series) -> tuple[pd.Series | None, str | None]:
    dt = pd.to_datetime(series, errors="coerce")
    if dt.isna().any():
        return None, "Failed to parse valid dates/timestamps in the first column."
    return dt, None


def _mask_hourly(dt: pd.Series, year: int) -> pd.Series:
    start = datetime(year, 1, 1, 1)
    end = datetime(year + 1, 1, 1, 1)
    return (dt >= start) & (dt < end)


def _mask_15m(dt: pd.Series, year: int) -> pd.Series:
    start = datetime(year, 1, 1, 0, 15)
    end = datetime(year + 1, 1, 1, 0, 15)
    return (dt >= start) & (dt < end)


def _explain(found: int, expected: list[int | str]) -> str:
    exp = ", ".join(str(e) for e in expected)
    return f"Found **{found}**, expected **{exp}**."


# ───────────────────────────── CORE DETECTOR ─────────────────────────────── #

def detect_format(df: pd.DataFrame) -> tuple[str, str]:
    """Return (*label*, *detail*).  *label* begins with "Error" on failure."""

    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    dt, err = _safe_to_datetime(df.iloc[:, 0])
    if dt is None:
        return "Error – date parsing", err  # type: ignore[return-value]

    has_times = (dt.dt.hour.gt(0) | dt.dt.minute.gt(0)).any()

    # ───────────────────────────── 1‑D PATH ──────────────────────────────── #
    if has_times:
        for y in sorted({d.year for d in dt}):
            # Hourly window
            rows_h = int(_mask_hourly(dt, y).sum())
            leap = calendar.isleap(y)
            exp_h_full = HOUR_ROWS_LEAP if leap else HOUR_ROWS_NONLEAP
            if rows_h in ACCEPTABLE_HOURLY:
                miss = " (final 00:00 missing)" if rows_h == exp_h_full - 1 else ""
                return (
                    "1D hourly",
                    f"{rows_h} rows for {y}{miss}. Expected {exp_h_full} rows starting 01:00.",
                )

            # 15‑minute window
            rows_q = int(_mask_15m(dt, y).sum())
            exp_q_full = Q15_ROWS_LEAP if leap else Q15_ROWS_NONLEAP
            if rows_q in ACCEPTABLE_15M:
                miss = " (final 00:00 missing)" if rows_q == exp_q_full - 1 else ""
                return (
                    "1D 15 minutes",
                    f"{rows_q} rows for {y}{miss}. Expected {exp_q_full} rows starting 00:15.",
                )

        # No acceptable year found
        summaries = [
            f"{y}: {_mask_hourly(dt, y).sum()}‑hour, {_mask_15m(dt, y).sum()}‑15‑min rows"
            for y in sorted({d.year for d in dt})
        ]
        return (
            "Error – no full 1‑D year",
            "Row counts → " + "; ".join(summaries),
        )  # type: ignore[return-value]

    # ───────────────────────────── 2‑D PATH ──────────────────────────────── #
    if (dt.dt.hour != 0).any() or (dt.dt.minute != 0).any():
        return (
            "Error – first column contains times",
            "2‑D diagrams must have *dates only* (midnight) in the first column.",
        )  # type: ignore[return-value]

    numeric_cols = [c for c in df.columns[1:] if pd.api.types.is_numeric_dtype(df[c])]
    n_num = len(numeric_cols)
    if n_num < 24:
        return "Error – too few interval columns", _explain(n_num, ["≥24"])  # type: ignore[return-value]

    gran = "hourly" if n_num < 96 else "15 minutes"
    expected_cols = 24 if gran == "hourly" else 96

    for y in sorted({d.year for d in dt}):
        rows = int((dt.dt.year == y).sum())
        if rows >= 365:
            extra = n_num - expected_cols
            extra_note = f" (+{extra} extra cols ignored)" if extra else ""
            return (
                f"2D {gran}",
                f"{rows} date‑rows for {y}. Using first {expected_cols} of {n_num} numeric columns{extra_note}.",
            )

    return (
        "Error – no full 2‑D year",
        "Sheet has dates but none of the years contains ≥365 rows.",
    )  # type: ignore[return-value]


# ───────────────────────────── STREAMLIT UI ──────────────────────────────── #

def main() -> None:  # noqa: D401
    st.set_page_config(page_title="Electricity Diagram Format Recognizer", page_icon="⚡")
    st.title("⚡ Electricity Diagram Format Recognizer")

    # Dependency check
    try:
        import openpyxl  # noqa: F401 – presence check
    except ModuleNotFoundError:
        st.error("Install **openpyxl** with `pip install openpyxl` and restart the app.")
        st.stop()

    upl = st.file_uploader("Upload an XLSX workbook", type=["xlsx", "xls"])
    if upl is None:
        st.info("⬆️ Drag‑and‑drop or browse to upload an Excel file.")
        st.stop()

    # File‑size guard
    if upl.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(f"File exceeds {MAX_UPLOAD_MB} MB upload limit. Please provide a smaller file.")
        st.stop()

    # Attempt to open workbook
    try:
        xls = pd.ExcelFile(upl)
    except InvalidFileException as exc:
        st.error("❌ The file appears to be corrupted or not a valid XLSX.")
        st.exception(exc)
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error("❌ Could not open Excel file.")
        st.exception(exc)
        st.stop()

    sheet = xls.sheet_names[0]
    if len(xls.sheet_names) > 1:
        sheet = st.selectbox("Select sheet", xls.sheet_names)

    # Parse selected sheet
    try:
        df = xls.parse(sheet)
    except Exception as exc:  # noqa: BLE001
        st.error("❌ Failed to parse the selected sheet.")
        st.exception(exc)
        st.stop()

    st.subheader("Data preview (first 5 rows, first 50 columns)")
    st.dataframe(df.iloc[:5, :50])

    # Detect format with safety net
    try:
        label, detail = detect_format(df)
    except Exception as exc:  # noqa: BLE001
        st.error("❌ Unexpected error while analysing the sheet.")
        st.exception(exc)
        st.stop()

    if label.startswith("Error"):
        st.error(label)
        st.write(detail)
    else:
        st.success(f"Detected format: **{label}**")
        st.write(detail)


if __name__ == "__main__":
    main()
