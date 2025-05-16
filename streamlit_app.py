import streamlit as st
import pandas as pd

# Ensure the correct Excel engine is available at runtime
try:
    import openpyxl  # noqa: F401
except ModuleNotFoundError:
    st.error(
        "The **openpyxl** package is required to read .xlsx files.\n\nInstall it via `pip install openpyxl` (or add `openpyxl` to your `requirements.txt` if you deploy to Streamlit Cloud) and restart the app."
    )
    st.stop()


def detect_format(df: pd.DataFrame) -> str:
    """Classify the Excel sheet structure.

    Returns one of:
        - "1D Diagram"
        - "2D Diagram"
        - "Error - unrecognized format"
    """
    # Basic sanitation: drop completely empty rows/cols and reset index
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    # 1D: two columns → timestamp + numeric consumption
    if df.shape[1] == 2:
        ts_col = df.iloc[:, 0]
        val_col = df.iloc[:, 1]
        try:
            pd.to_datetime(ts_col)
            if pd.api.types.is_numeric_dtype(val_col):
                return "1D Diagram"
        except Exception:  # noqa: BLE001 – parsing failure means not 1D
            pass

    # 2D: ≥3 columns → date + multiple interval columns (numeric)
    if df.shape[1] >= 3:
        date_col = df.iloc[:, 0]
        interval_df = df.iloc[:, 1:]
        try:
            pd.to_datetime(date_col)
            if all(pd.api.types.is_numeric_dtype(interval_df[c]) for c in interval_df.columns):
                return "2D Diagram"
        except Exception:  # noqa: BLE001
            pass

    return "Error - unrecognized format"


def main() -> None:
    st.set_page_config(page_title="Electricity Diagram Format Recognizer", page_icon="⚡")
    st.title("⚡ Electricity Diagram Format Recognizer")

    uploaded_file = st.file_uploader(
        "Upload an XLSX file containing your consumption diagram",
        type=["xlsx", "xls"],
    )

    if uploaded_file is None:
        st.info("⬆️  Drag & drop or browse to upload an Excel file.")
        st.stop()

    try:
        workbook = pd.ExcelFile(uploaded_file)
        sheet_names = workbook.sheet_names
        sheet = sheet_names[0]
        if len(sheet_names) > 1:
            sheet = st.selectbox("Select sheet to analyze", sheet_names)

        df = workbook.parse(sheet)
        st.subheader("Data preview (first 5 rows)")
        st.dataframe(df.head())

        detected = detect_format(df)
        st.success(f"Detected format: **{detected}**")

    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Could not read the Excel file: {exc}")


if __name__ == "__main__":
    main()
