# ============================================================
# MR Linelist: map every variable using the Mapping sheet
# Run in Google Colab (paste into a cell and run)
# ============================================================
import io
import warnings
import numpy as np
import pandas as pd
from google.colab import files

warnings.filterwarnings("ignore")

MASTER_SHEET = "Master List Data"
MAPPING_SHEET = "Mapping"
OUTPUT_NAME = "MR_Linelist_Mapped.xlsx"

# ---- 1) Upload the file --------------------------------------
print("Choose MR_Linelist.xlsx to upload...")
uploaded = files.upload()
if not uploaded:
    raise SystemExit("No file uploaded.")
file_name = next(iter(uploaded))
print(f"Uploaded: {file_name}")
xls = pd.ExcelFile(io.BytesIO(uploaded[file_name]))

# ---- 2) Read master data and mapping table -------------------
master = pd.read_excel(xls, sheet_name=MASTER_SHEET, dtype=object)

# Mapping sheet: column C = variable, D = raw input, E = mapped value
mp = pd.read_excel(xls, sheet_name=MAPPING_SHEET, usecols="C:E", header=0, dtype=object)
mp.columns = ["var", "input", "mapping"]
mp = mp.dropna(subset=["var", "mapping"])  # keep only rows that have a mapping


def norm(v):
    """Make raw values comparable: trim, lowercase, 10.0 -> '10'."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (int, float)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip().lower()


lookup = {}
for var, inp, m in zip(mp["var"], mp["input"], mp["mapping"]):
    lookup.setdefault((var, norm(inp)), m)  # first mapping wins

mapped_vars = {k[0] for k in lookup}

# ---- 3) Apply mapping (fallback = keep original value) -------
mapped = master.copy()
report = []
for col in master.columns:
    if col not in mapped_vars:
        continue  # variable has no mapping -> leave untouched
    new_vals = [
        lookup.get((col, norm(v)), v) if norm(v) is not None else v
        for v in master[col]
    ]
    changed = sum(
        1 for a, b in zip(master[col], new_vals)
        if a != b and not (pd.isna(a) and pd.isna(b))
    )
    mapped[col] = new_vals
    report.append((col, changed))

print(pd.DataFrame(report, columns=["variable", "cells_changed"]).to_string())
print("Rows x columns:", mapped.shape)

# ---- 4) Save and download ------------------------------------
with pd.ExcelWriter(OUTPUT_NAME, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
    mapped.to_excel(xw, sheet_name="Mapped Data", index=False)

files.download(OUTPUT_NAME)
print("Done. Your browser should download", OUTPUT_NAME)
