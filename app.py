import io
import warnings

import numpy as np
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

MASTER_SHEET = "Master List Data"
MAPPING_SHEET = "Mapping"

st.set_page_config(page_title="MR Linelist Mapper", layout="wide")
st.title("MR Linelist Mapper")
st.write(
    "Upload the MR Linelist workbook. Every variable that has a mapping in the "
    f"**{MAPPING_SHEET}** sheet is converted; all other columns are kept as they are."
)


def norm(v):
    """Make raw values comparable: trim, lowercase, 10.0 -> '10'."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (int, float)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip().lower()


def process(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    missing = [s for s in (MASTER_SHEET, MAPPING_SHEET) if s not in xls.sheet_names]
    if missing:
        raise ValueError(f"Sheet(s) not found: {missing}. Found: {xls.sheet_names}")

    master = pd.read_excel(xls, sheet_name=MASTER_SHEET, dtype=object)

    # Mapping sheet: column C = variable, D = raw input, E = mapped value
    mp = pd.read_excel(xls, sheet_name=MAPPING_SHEET, usecols="C:E", header=0, dtype=object)
    mp.columns = ["var", "input", "mapping"]
    mp = mp.dropna(subset=["var", "mapping"])

    lookup = {}
    for var, inp, m in zip(mp["var"], mp["input"], mp["mapping"]):
        lookup.setdefault((var, norm(inp)), m)  # first mapping wins
    mapped_vars = {k[0] for k in lookup}

    mapped = master.copy()
    report = []
    for col in master.columns:
        if col not in mapped_vars:
            continue
        new_vals = [
            lookup.get((col, norm(v)), v) if norm(v) is not None else v
            for v in master[col]
        ]
        changed = sum(
            1 for a, b in zip(master[col], new_vals)
            if a != b and not (pd.isna(a) and pd.isna(b))
        )
        mapped[col] = new_vals
        report.append({"variable": col, "cells_changed": changed})

    return master, mapped, pd.DataFrame(report)


def to_xlsx(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, sheet_name="Mapped Data", index=False)
    return buf.getvalue()


uploaded = st.file_uploader("Upload MR_Linelist.xlsx", type=["xlsx"])

if uploaded is not None:
    try:
        with st.spinner("Mapping values... large files can take a minute or two."):
            master, mapped, report = process(uploaded.getvalue())
    except Exception as e:
        st.error(f"Could not process the file: {e}")
        st.stop()

    st.success(f"Done: {mapped.shape[0]:,} rows x {mapped.shape[1]} columns.")

    st.subheader("Cells changed per variable")
    st.dataframe(report, use_container_width=True, hide_index=True)

    st.subheader("Preview (first 100 rows)")
    st.dataframe(mapped.head(100).astype(str), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        with st.spinner("Preparing Excel file..."):
            xlsx_bytes = to_xlsx(mapped)
        st.download_button(
            "Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="MR_Linelist_Mapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        st.download_button(
            "Download CSV (faster, smaller)",
            data=mapped.to_csv(index=False).encode("utf-8-sig"),
            file_name="MR_Linelist_Mapped.csv",
            mime="text/csv",
        )
