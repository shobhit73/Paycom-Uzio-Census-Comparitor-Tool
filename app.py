# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom vs UZIO – Census Audit Tool
# INPUT workbook tabs (single .xlsx upload):
#   - Uzio Data
#   - Paycom Data
#   - Mapping Sheet
#
# OUTPUT workbook tabs:
#   - Summary
#   - Field_Summary_By_Status
#   - Comparison_Detail_AllFields
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET = "Uzio Data"
PAYCOM_SHEET = "Paycom Data"
MAP_SHEET = "Mapping Sheet"

# ---------- UI ----------
st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
st.markdown(
    """
    <style>
      [data-testid="stSidebar"] { display: none !important; }
      [data-testid="collapsedControl"] { display: none !important; }
      header { display: none !important; }
      footer { display: none !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Helpers ----------
def norm_colname(c: str) -> str:
    if c is None:
        return ""
    c = str(c).replace("\n", " ").replace("\r", " ")
    c = c.replace("\u00A0", " ")
    c = c.replace("’", "'").replace("“", '"').replace("”", '"')
    c = re.sub(r"\s+", " ", c).strip()
    c = c.replace("*", "")
    c = c.strip('"').strip("'")
    return c

def norm_blank(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    if isinstance(x, str) and x.strip().lower() in {"", "nan", "none", "null"}:
        return ""
    return x

def digits_only(x):
    """Extract digits while handling numeric types safely (avoid '.0' artifacts)."""
    x = norm_blank(x)
    if x == "":
        return ""
    try:
        if isinstance(x, (int, np.integer)):
            return str(int(x))
        if isinstance(x, (float, np.floating)):
            if float(x).is_integer():
                return str(int(x))
    except Exception:
        pass

    s = str(x).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return re.sub(r"\D", "", s)

def digits_only_padded(x, width: int):
    d = digits_only(x)
    if d == "":
        return ""
    if len(d) < width:
        d = d.zfill(width)
    return d

def try_parse_date(x):
    """
    Normalize dates so these match:
      - 07/15/2024
      - 2024-07-15 00:00:00
      - datetime objects
    Output as ISO yyyy-mm-dd when parseable.
    """
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x).date().isoformat()

    s = str(x).strip()
    if s == "":
        return ""
    # common Paycom placeholder
    if s in {"00/00/0000", "0/0/0000", "0000-00-00"}:
        return ""
    try:
        return pd.to_datetime(s, errors="raise").date().isoformat()
    except Exception:
        return s

def norm_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

def find_col(df_cols, *candidate_names):
    norm_map = {norm_colname(c).casefold(): c for c in df_cols}
    for cand in candidate_names:
        key = norm_colname(cand).casefold()
        if key in norm_map:
            return norm_map[key]
    return None

def resolve_paycom_col_label(label: str, paycom_cols_all) -> str:
    """
    Mapping sheet may contain values with commas/quotes/etc.
    Resolve to actual Paycom column names present in the sheet.
    """
    if label is None:
        return ""
    raw = str(label).strip()
    raw = raw.replace("’", "'").replace("“", '"').replace("”", '"')
    raw = raw.strip().strip(",")
    if raw == "":
        return ""

    pay_norm = {norm_colname(c).casefold(): c for c in paycom_cols_all}

    direct = norm_colname(raw).casefold()
    if direct in pay_norm:
        return pay_norm[direct]

    parts = re.split(r"\(|\)|\bor\b|/|,|;", raw, flags=re.IGNORECASE)
    parts = [norm_colname(p) for p in parts if norm_colname(p)]
    for p in parts:
        k = norm_colname(p).casefold()
        if k in pay_norm:
            return pay_norm[k]

    for k_norm, actual in pay_norm.items():
        if k_norm and (k_norm in direct or direct in k_norm):
            return actual

    return ""

# ---------- Field normalization ----------
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
PHONE_KEYWORDS = {"phone", "mobile"}
DATE_KEYWORDS = {"date", "dob", "birth", "effective", "doh", "hire"}
SSN_KEYWORDS = {"ssn", "social"}
SUFFIX_KEYWORDS = {"suffix"}
PAYTYPE_KEYWORDS = {"pay type"}
EMP_STATUS_KEYWORDS = {"employment status"}
TERMINATION_REASON_KEYWORDS = {"termination reason"}
HOURLY_RATE_KEYWORDS = {"hourly rate", "rate per hour", "hourlyrate"}

def norm_phone_digits(x):
    d = digits_only(x)
    if d == "":
        return ""
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) > 10:
        d = d[-10:]
    return d

def norm_zip_first5(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)) and float(x).is_integer():
        s = str(int(x))
    else:
        s = re.sub(r"[^\d]", "", str(x).strip())
    if s == "":
        return ""
    if 0 < len(s) < 5:
        s = s.zfill(5)
    return s[:5]

def norm_suffix(x):
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).strip()
    s = s.replace(".", "").replace(",", "")
    s = re.sub(r"\s+", "", s)
    return s.casefold()

def norm_pay_type(x):
    s = norm_blank(x)
    if s == "":
        return ""
    s = re.sub(r"\s+", " ", str(s).strip()).casefold()
    if s in {"salaried", "salary"}:
        return "salary"
    if s in {"hourly", "hour"}:
        return "hourly"
    return s

def norm_employment_status(x):
    """
    Business rule: Paycom 'On Leave' should be considered Active in UZIO.
    """
    s = norm_blank(x)
    if s == "":
        return ""
    s = re.sub(r"\s+", " ", str(s).strip()).casefold()
    if "on leave" in s:
        return "active"
    if s == "active":
        return "active"
    if s in {"terminated", "inactive"}:
        return s
    return s

def norm_termination_reason(uzio_val, paycom_val, side: str):
    """
    Rules:
      - If Paycom contains 'voluntary' and UZIO contains 'voluntary' anywhere => match.
      - If Paycom contains 'involuntary' and UZIO contains 'involuntary' anywhere => match.
      - Otherwise, Paycom reason buckets to 'other' and matches UZIO 'other'.
    """
    if side not in {"uzio", "paycom"}:
        side = "uzio"

    u = re.sub(r"\s+", " ", str(norm_blank(uzio_val)).strip()).casefold()
    p = re.sub(r"\s+", " ", str(norm_blank(paycom_val)).strip()).casefold()

    if "involuntary" in p:
        return "involuntary" if side == "paycom" else ("involuntary" if "involuntary" in u else u)
    if "voluntary" in p:
        return "voluntary" if side == "paycom" else ("voluntary" if "voluntary" in u else u)

    if side == "paycom":
        return "other"
    return "other" if u == "other" else u

def norm_value(x, field_name: str):
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    if any(k in f for k in SSN_KEYWORDS):
        return digits_only_padded(x, 9)

    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)

    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    if any(k in f for k in SUFFIX_KEYWORDS):
        return norm_suffix(x)

    if any(k in f for k in PAYTYPE_KEYWORDS):
        return norm_pay_type(x)

    if any(k in f for k in EMP_STATUS_KEYWORDS):
        return norm_employment_status(x)

    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()
    return str(x).casefold()

def read_mapping_sheet(xls: pd.ExcelFile, sheet_name: str, paycom_cols_all: list) -> pd.DataFrame:
    m = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
    m.columns = [norm_colname(c) for c in m.columns]

    uz_col_name = None
    pay_col_name = None
    for c in m.columns:
        cc = norm_colname(c).casefold()
        if cc in {"uzio coloumn", "uzio column"}:
            uz_col_name = c
        if cc in {"paycom coloumn", "paycom column"}:
            pay_col_name = c

    if uz_col_name is None or pay_col_name is None:
        raise ValueError(f"'{sheet_name}' must contain columns: 'UZIO Column' and 'Paycom Column'.")

    m[uz_col_name] = m[uz_col_name].map(norm_colname)
    m[pay_col_name] = m[pay_col_name].map(norm_colname)

    m = m.dropna(subset=[uz_col_name, pay_col_name]).copy()
    m = m[(m[uz_col_name] != "") & (m[pay_col_name] != "")]
    m = m.drop_duplicates(subset=[uz_col_name], keep="first").copy()

    m["UZIO_Column"] = m[uz_col_name]
    m["PAYCOM_Label"] = m[pay_col_name]
    m["PAYCOM_Resolved_Column"] = m["PAYCOM_Label"].map(lambda x: resolve_paycom_col_label(x, paycom_cols_all))

    m["_uz_norm"] = m["UZIO_Column"].map(lambda x: norm_colname(x).casefold())
    m = m[m["_uz_norm"] != "employee id"].copy()
    m.drop(columns=["_uz_norm"], inplace=True)

    return m

# ---------- Core comparison ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio = pd.read_excel(xls, sheet_name=UZIO_SHEET, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=PAYCOM_SHEET, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]

    UZIO_KEY = find_col(uzio.columns, "Employee ID", "EmployeeID", "Employee Id")
    if UZIO_KEY is None:
        raise ValueError("UZIO key column not found (expected 'Employee ID') in 'Uzio Data'.")

    # ✅ FIX: Paycom key supports Employee_Code
    PAYCOM_KEY = find_col(
        paycom.columns,
        "Employee_Code", "Employee Code", "Employee ID", "EmployeeID", "Employee", "Employee Id",
    )
    if PAYCOM_KEY is None:
        raise ValueError(
            "Paycom key column not found (expected 'Employee_Code' or 'Employee ID' or 'Employee') in 'Paycom Data'."
        )

    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    mapping = read_mapping_sheet(xls, MAP_SHEET, list(paycom.columns))
    sec_map = mapping[mapping["PAYCOM_Resolved_Column"] != ""].copy()

    UZIO_EMP_STATUS_COL = find_col(uzio.columns, "Employment Status")
    UZIO_PAYTYPE_COL = find_col(uzio.columns, "Pay Type")

    uzio_first = uzio.groupby(UZIO_KEY, sort=False).head(1).set_index(UZIO_KEY)
    paycom_first = paycom.groupby(PAYCOM_KEY, sort=False).head(1).set_index(PAYCOM_KEY)

    all_emp_ids = sorted(set(uzio_first.index.astype(str)) | set(paycom_first.index.astype(str)))

    rows = []
    for emp_id in all_emp_ids:
        emp_id = str(emp_id)

        uz_exists = emp_id in uzio_first.index
        pay_exists = emp_id in paycom_first.index

        emp_status_val = ""
        if uz_exists and UZIO_EMP_STATUS_COL and UZIO_EMP_STATUS_COL in uzio_first.columns:
            emp_status_val = uzio_first.loc[emp_id, UZIO_EMP_STATUS_COL]

        uzio_pay_type_val = ""
        if uz_exists and UZIO_PAYTYPE_COL and UZIO_PAYTYPE_COL in uzio_first.columns:
            uzio_pay_type_val = uzio_first.loc[emp_id, UZIO_PAYTYPE_COL]
        uzio_pay_type_norm = norm_pay_type(uzio_pay_type_val)

        for _, r in sec_map.iterrows():
            uz_field = r["UZIO_Column"]
            pay_col = r["PAYCOM_Resolved_Column"]

            f_norm = norm_colname(uz_field).casefold()

            # Salaried => skip hourly rate comparisons
            if uzio_pay_type_norm == "salary":
                if any(k in f_norm for k in HOURLY_RATE_KEYWORDS):
                    continue

            uz_val = uzio_first.loc[emp_id, uz_field] if (uz_exists and uz_field in uzio_first.columns) else ""
            pay_val = paycom_first.loc[emp_id, pay_col] if (pay_exists and pay_col in paycom_first.columns) else ""

            if (not pay_exists) and uz_exists:
                status = "MISSING_IN_PAYCOM"
            elif pay_exists and (not uz_exists):
                status = "MISSING_IN_UZIO"
            elif pay_col not in paycom.columns:
                status = "PAYCOM_COLUMN_MISSING"
            elif uz_field not in uzio.columns:
                status = "UZIO_COLUMN_MISSING"
            else:
                if any(k in f_norm for k in TERMINATION_REASON_KEYWORDS):
                    uz_n = norm_termination_reason(uz_val, pay_val, side="uzio")
                    pay_n = norm_termination_reason(uz_val, pay_val, side="paycom")
                else:
                    uz_n = norm_value(uz_val, uz_field)
                    pay_n = norm_value(pay_val, uz_field)

                if (uz_n == pay_n) or (uz_n == "" and pay_n == ""):
                    status = "OK"
                elif uz_n == "" and pay_n != "":
                    status = "UZIO_MISSING_VALUE"
                elif uz_n != "" and pay_n == "":
                    status = "PAYCOM_MISSING_VALUE"
                else:
                    status = "MISMATCH"

            rows.append(
                {
                    "Employee ID": emp_id,
                    "Field": uz_field,
                    "Employment Status": emp_status_val,  # ✅ extra context column after Field
                    "UZIO_Value": uz_val,
                    "PAYCOM_Value": pay_val,
                    "PAYCOM_SourceOfTruth_Status": status,
                }
            )

    comparison_detail = pd.DataFrame(
        rows,
        columns=[
            "Employee ID",
            "Field",
            "Employment Status",
            "UZIO_Value",
            "PAYCOM_Value",
            "PAYCOM_SourceOfTruth_Status",
        ],
    )

    statuses = [
        "OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "PAYCOM_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_PAYCOM",
        "PAYCOM_COLUMN_MISSING",
        "UZIO_COLUMN_MISSING",
    ]

    if len(comparison_detail):
        field_summary_by_status = (
            comparison_detail.pivot_table(
                index="Field",
                columns="PAYCOM_SourceOfTruth_Status",
                values="Employee ID",
                aggfunc="count",
                fill_value=0,
            )
            .reindex(columns=statuses, fill_value=0)
            .reset_index()
        )
        field_summary_by_status["Total"] = field_summary_by_status[statuses].sum(axis=1)
    else:
        field_summary_by_status = pd.DataFrame(columns=["Field"] + statuses + ["Total"])

    uzio_emp = set(uzio_first.index.astype(str))
    paycom_emp = set(paycom_first.index.astype(str))

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total UZIO Employees",
                "Total Paycom Employees",
                "Employees in both",
                "Employees only in UZIO",
                "Employees only in Paycom",
                "Fields Compared",
                "Total Comparisons (field-level rows)",
            ],
            "Value": [
                len(uzio_emp),
                len(paycom_emp),
                len(uzio_emp & paycom_emp),
                len(uzio_emp - paycom_emp),
                len(paycom_emp - uzio_emp),
                int(sec_map.shape[0]),
                int(comparison_detail.shape[0]),
            ],
        }
    )

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Field_Summary_By_Status", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison_Detail_AllFields", index=False)

    return out.getvalue()

# ---------- UI ----------
st.title(APP_TITLE)
st.write("Upload the Excel workbook (.xlsx) with 3 tabs: Uzio Data, Paycom Data, and Mapping Sheet.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
run_btn = st.button("Run Audit", type="primary", disabled=(uploaded_file is None))

if run_btn:
    try:
        with st.spinner("Running audit..."):
            report_bytes = run_comparison(uploaded_file.getvalue())

        st.success("Report generated.")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_name = f"UZIO_vs_PAYCOM_Comparison_Report_PAYCOM_SourceOfTruth_{ts}.xlsx"

        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=report_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
