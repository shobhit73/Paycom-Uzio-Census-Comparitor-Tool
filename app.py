# ============================
# Paycom vs UZIO Census Audit Tool
# ============================

import io
import re
from datetime import datetime, date
from decimal import Decimal, InvalidOperation

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET = "Uzio Data"
PAYCOM_SHEET = "Paycom Data"
MAP_SHEET = "Mapping Sheet"

# ---------- UI ----------
st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
st.markdown("""
<style>
[data-testid="stSidebar"], header, footer { display:none !important; }
</style>
""", unsafe_allow_html=True)

# ---------- Helpers ----------
def norm_colname(c):
    if c is None:
        return ""
    c = str(c).replace("\n", " ").replace("\r", " ")
    c = c.replace("\u00A0", " ")
    c = re.sub(r"\s+", " ", c).strip()
    return c

def norm_blank(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    s = str(x).strip()
    return "" if s.lower() in {"", "nan", "null", "none"} else s

def norm_key_series(s):
    return s.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

def digits_only(x):
    return re.sub(r"\D", "", str(x)) if x else ""

def norm_phone(x):
    d = digits_only(x)
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    return d[-10:] if len(d) >= 10 else d

def norm_zip(x):
    d = digits_only(x)
    return d[:5]

def safe_decimal(x):
    try:
        return Decimal(str(x).replace(",", ""))
    except:
        return None

def approx_equal(a, b):
    return a is not None and b is not None and a == b

# ---------- Mapping Sheet ----------
def read_mapping_sheet(xls, sheet, uz_cols, pc_cols):
    m = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    m.columns = [norm_colname(c) for c in m.columns]

    uz_col = "UZIO Column"
    pc_col = "Paycom Column"

    m[uz_col] = m[uz_col].map(norm_colname)
    m[pc_col] = m[pc_col].map(norm_colname)

    m["_uz_norm"] = m[uz_col].str.casefold()

    # ✅ FIXED: pandas-safe filter
    m = m[~m["_uz_norm"].isin(["employee id", "employee", "employee_code", "employee code"])].copy()

    def resolve(col, cols):
        col = col.casefold()
        for c in cols:
            if col == c.casefold():
                return c
        return ""

    m["PAYCOM_Resolved"] = m[pc_col].apply(lambda x: resolve(x, pc_cols))
    return m

# ---------- Core ----------
def run_comparison(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio = pd.read_excel(xls, UZIO_SHEET, dtype=object)
    paycom = pd.read_excel(xls, PAYCOM_SHEET, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]

    UZ_KEY = next(c for c in uzio.columns if "employee" in c.lower())
    PC_KEY = next(c for c in paycom.columns if "employee" in c.lower())

    uzio[UZ_KEY] = norm_key_series(uzio[UZ_KEY])
    paycom[PC_KEY] = norm_key_series(paycom[PC_KEY])

    mapping = read_mapping_sheet(xls, MAP_SHEET, uzio.columns, paycom.columns)

    uz_by = uzio.set_index(UZ_KEY)
    pc_by = paycom.set_index(PC_KEY)

    employees = sorted(set(uz_by.index) | set(pc_by.index))

    rows = []

    for emp in employees:
        uz = uz_by.loc[emp] if emp in uz_by.index else None
        pc = pc_by.loc[emp] if emp in pc_by.index else None

        if isinstance(uz, pd.DataFrame):
            uz = uz.iloc[0]
        if isinstance(pc, pd.DataFrame):
            pc = pc.iloc[0]

        emp_status = uz.get("Employment Status", "") if uz is not None else ""

        pay_type = uz.get("Pay Type", "") if uz is not None else ""

        for _, r in mapping.iterrows():
            field = r["UZIO Column"]
            pc_col = r["PAYCOM_Resolved"]

            uz_val = uz.get(field, "") if uz is not None else ""
            pc_val = pc.get(pc_col, "") if pc is not None and pc_col else ""

            status = "OK"

            fname = field.lower()

            if "phone" in fname:
                status = "OK" if norm_phone(uz_val) == norm_phone(pc_val) else "MISMATCH"

            elif "zipcode" in fname:
                status = "OK" if norm_zip(uz_val) == norm_zip(pc_val) else "MISMATCH"

            elif "middle initial" in fname:
                status = "OK" if (str(uz_val)[:1].upper() == str(pc_val)[:1].upper()) else "MISMATCH"

            elif "annual salary" in fname:
                if "hour" in pay_type.lower():
                    status = "OK"
                else:
                    status = "OK" if approx_equal(safe_decimal(uz_val), safe_decimal(pc_val)) else "MISMATCH"

            elif "hourly pay rate" in fname or "working hours" in fname:
                if "salary" in pay_type.lower():
                    status = "OK"
                else:
                    status = "OK" if approx_equal(safe_decimal(uz_val), safe_decimal(pc_val)) else "MISMATCH"

            rows.append({
                "Employee": emp,
                "Field": field,
                "Employment Status": emp_status,
                "UZIO_Value": uz_val,
                "PAYCOM_Value": pc_val,
                "PAYCOM_SourceOfTruth_Status": status
            })

    df = pd.DataFrame(rows)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, "Comparison_Detail_AllFields", index=False)

    return out.getvalue()

# ---------- UI ----------
st.title(APP_TITLE)
file = st.file_uploader("Upload Excel workbook (.xlsx)", type=["xlsx"])

if st.button("Run Audit", disabled=file is None):
    try:
        data = run_comparison(file.getvalue())
        st.download_button(
            "Download Report",
            data,
            f"UZIO_vs_PAYCOM_Comparison_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Failed: {e}")
