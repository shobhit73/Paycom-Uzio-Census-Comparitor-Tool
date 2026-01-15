# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom Uzio Census Audit Tool
#
# INPUT workbook tabs (single .xlsx upload):
#   - Uzio Data
#   - Paycom Data
#   - Mapping Sheet   (or "Mapping")
#
# OUTPUT workbook tabs:
#   - Summary
#   - Field_Summary_By_Status
#   - Comparison_Detail_AllFields
#
# CHANGE REQUESTS IMPLEMENTED (ONLY):
#   1) Add "Employment Status" column right after "Field" in Comparison_Detail_AllFields,
#      populated from UZIO Employment Status (mandatory in UZIO).
#   2) Employment Status comparison: PAYCOM "On Leave" should be treated as UZIO "Active" (not a mismatch).
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET = "Uzio Data"
PAYCOM_SHEET = "Paycom Data"
MAPPING_SHEET_CANDIDATES = ["Mapping Sheet", "Mapping"]

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
    unsafe_allow_html=True
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
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x).date().isoformat()
    if isinstance(x, str):
        s = x.strip()
        try:
            return pd.to_datetime(s, errors="raise").date().isoformat()
        except Exception:
            return s
    return str(x)

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

def norm_phone_digits(x):
    d = digits_only(x)
    if d == "":
        return ""
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) > 10:
        d = d[-10:]
    return d

def normalize_suffix(x):
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).strip()
    s = re.sub(r"[^A-Za-z0-9]", "", s)  # remove dots/spaces
    return s.casefold()

def normalize_middle_initial(x):
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).strip()
    # take first alpha-numeric character as initial
    m = re.search(r"[A-Za-z0-9]", s)
    return (m.group(0) if m else "").casefold()

def find_col(df_cols, *candidate_names):
    norm_map = {norm_colname(c).casefold(): c for c in df_cols}
    for cand in candidate_names:
        key = norm_colname(cand).casefold()
        if key in norm_map:
            return norm_map[key]
    return None

def resolve_mapping_sheet_name(xls: pd.ExcelFile) -> str:
    sheets = set(xls.sheet_names)
    for s in MAPPING_SHEET_CANDIDATES:
        if s in sheets:
            return s
    raise ValueError(f"Mapping sheet not found. Expected one of: {MAPPING_SHEET_CANDIDATES}")

def read_mapping_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    m = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
    m.columns = [norm_colname(c) for c in m.columns]

    uz_col_name = None
    pc_col_name = None
    for c in m.columns:
        if norm_colname(c).casefold() in {"uzio coloumn", "uzio column"}:
            uz_col_name = c
        if norm_colname(c).casefold() in {"paycom coloumn", "paycom column"}:
            pc_col_name = c

    if uz_col_name is None or pc_col_name is None:
        raise ValueError(f"'{sheet_name}' must contain columns: 'UZIO Column' and 'Paycom Column'.")

    m[uz_col_name] = m[uz_col_name].map(norm_colname)
    m[pc_col_name] = m[pc_col_name].map(norm_colname)

    m = m.dropna(subset=[uz_col_name, pc_col_name]).copy()
    m = m[(m[uz_col_name] != "") & (m[pc_col_name] != "")]
    m = m.drop_duplicates(subset=[uz_col_name], keep="first").copy()

    m["UZIO_Column"] = m[uz_col_name]
    m["PAYCOM_Column"] = m[pc_col_name]

    # exclude Employee ID row from comparisons (it is only key)
    m["_uz_norm"] = m["UZIO_Column"].map(lambda x: norm_colname(x).casefold())
    m = m[m["_uz_norm"] not in ["employee id"]].copy() if len(m) else m
    if "_uz_norm" in m.columns:
        m.drop(columns=["_uz_norm"], inplace=True, errors="ignore")

    return m

# Keywords
PHONE_KEYWORDS = {"phone", "mobile"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
DATE_KEYWORDS = {"date", "dob", "birth", "effective", "doh", "hire"}
SSN_KEYWORDS = {"ssn", "social security"}
MIDDLE_INITIAL_KEYWORDS = {"middle initial"}
SUFFIX_KEYWORDS = {"suffix"}
PAYTYPE_KEYWORDS = {"pay type"}
EMP_STATUS_KEYWORDS = {"employment status"}
TERMINATION_REASON_KEYWORDS = {"termination reason"}

# Pay type synonyms
PAYTYPE_SYNONYMS = {
    "salary": "salaried",
    "salaried": "salaried",
    "hourly": "hourly",
}

def norm_value(x, field_name: str):
    """
    Normalize values for comparison.

    NOTE: Only special-case added here is:
      - Employment Status: treat Paycom "On Leave" as "Active"
    """
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    # Phone normalization
    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)

    # Zip normalization
    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    # SSN normalization (preserve leading zeros)
    if any(k in f for k in SSN_KEYWORDS):
        return digits_only_padded(x, 9)

    # Middle Initial: compare first character only (Nicole vs N)
    if any(k in f for k in MIDDLE_INITIAL_KEYWORDS):
        return normalize_middle_initial(x)

    # Suffix: Jr. vs JR
    if any(k in f for k in SUFFIX_KEYWORDS):
        return normalize_suffix(x)

    # Dates (fix Original DOH time vs date)
    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    # Pay Type: Salaried == Salary
    if any(k in f for k in PAYTYPE_KEYWORDS):
        s = re.sub(r"\s+", " ", str(x).strip()).casefold()
        return PAYTYPE_SYNONYMS.get(s, s)

    # Employment Status (CHANGE): Paycom "On Leave" == Uzio "Active"
    if any(k in f for k in EMP_STATUS_KEYWORDS):
        s = re.sub(r"\s+", " ", str(x).strip()).casefold()
        if s == "on leave":
            s = "active"
        return s

    # default: string compare (case-insensitive, collapse spaces)
    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()

    return str(x).casefold()

def key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

# Termination Reason mapping (Paycom -> allowed Uzio values)
# (kept flexible: if Paycom value matches, Uzio can be any allowed value)
TERM_REASON_ALLOWED = {
    "attendance violation": {"other"},
    "employee elected not to return": {"other"},
    "failure to show up for work": {"other"},
    "poor performance and attendance": {"other"},
    "attendance violation and poor performance": {"other"},
    "never worked": {"other"},
    "performance, safety, and attendance violations": {"other"},
    "left for military service": {"other"},
    "policy violation": {"other"},
    "personal issue": {"other"},
    "poor performance, poor attitude and attendance violations": {"other"},
    "health issue": {"other"},
    "conflict with other job": {"other"},
    "poor performance": {"other"},
    "did not start work": {"other"},
    "poor attitude and performance": {"other"},
    "poor attendance, attitude, and performance.": {"other"},
    "professionalism, performance and attendance issues": {"other"},
    "voluntary termination of employment": {"voluntary termination", "voluntary resignation", "voluntary"},
    "involuntary termination of employment": {"involuntary termination"},
}

def termination_reason_ok(uz_raw, pc_raw) -> bool:
    uz = re.sub(r"\s+", " ", str(norm_blank(uz_raw) or "").strip()).casefold()
    pc = re.sub(r"\s+", " ", str(norm_blank(pc_raw) or "").strip()).casefold()

    if uz == "" and pc == "":
        return True

    # If both contain "voluntary" (and not "involuntary"), OK
    if "voluntary" in uz and "voluntary" in pc and ("involuntary" not in uz) and ("involuntary" not in pc):
        return True

    # If both contain "involuntary", OK
    if "involuntary" in uz and "involuntary" in pc:
        return True

    # Mapping table (Paycom -> allowed Uzio)
    allowed = TERM_REASON_ALLOWED.get(pc, None)
    if allowed is not None:
        return uz in allowed

    return False

def get_pay_type_for_employee(emp_id: str, uzio_df: pd.DataFrame, paycom_df: pd.DataFrame, uz_key: str, pc_key: str):
    # Prefer UZIO Pay Type
    uz_pay_col = find_col(uzio_df.columns, "Pay Type")
    if uz_pay_col and emp_id in uzio_df.index:
        v = uzio_df.loc[emp_id, uz_pay_col]
        if isinstance(v, pd.Series):
            v = v.iloc[0]
        s = str(norm_blank(v) or "").strip().casefold()
        if s:
            return PAYTYPE_SYNONYMS.get(s, s)

    # Fallback Paycom Pay Type
    pc_pay_col = find_col(paycom_df.columns, "Pay Type")
    if pc_pay_col and emp_id in paycom_df.index:
        v = paycom_df.loc[emp_id, pc_pay_col]
        if isinstance(v, pd.Series):
            v = v.iloc[0]
        s = str(norm_blank(v) or "").strip().casefold()
        if s:
            return PAYTYPE_SYNONYMS.get(s, s)

    return ""

def should_skip_field_for_employee(field_name: str, pay_type_norm: str) -> bool:
    """
    Existing behavior requested earlier:
      - If person is salaried, don't compare their hourly rate.
    """
    f = norm_colname(field_name).casefold()
    if pay_type_norm == "salaried":
        if "hourly rate" in f or ("rate" in f and "hour" in f):
            return True
    return False

def build_report(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    mapping_sheet = resolve_mapping_sheet_name(xls)

    uzio = pd.read_excel(xls, sheet_name=UZIO_SHEET, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=PAYCOM_SHEET, dtype=object)
    mapping = read_mapping_sheet(xls, mapping_sheet)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]

    # Key columns
    uz_key = find_col(uzio.columns, "Employee ID", "Employee", "EmployeeID", "Emp ID")
    pc_key = find_col(paycom.columns, "Employee ID", "Employee", "EmployeeID", "Emp ID")
    if uz_key is None:
        raise ValueError("UZIO key column not found (expected 'Employee ID' or 'Employee') in 'Uzio Data'.")
    if pc_key is None:
        raise ValueError("Paycom key column not found (expected 'Employee ID' or 'Employee') in 'Paycom Data'.")

    # Normalize keys
    uzio[uz_key] = key_series(uzio[uz_key])
    paycom[pc_key] = key_series(paycom[pc_key])

    # Index by Employee (keep first if duplicates)
    uzio_idx = uzio.dropna(subset=[uz_key]).copy()
    paycom_idx = paycom.dropna(subset=[pc_key]).copy()
    uzio_idx = uzio_idx[uzio_idx[uz_key].map(norm_blank) != ""]
    paycom_idx = paycom_idx[paycom_idx[pc_key].map(norm_blank) != ""]

    uzio_idx = uzio_idx.groupby(uz_key, as_index=False).first().set_index(uz_key, drop=False)
    paycom_idx = paycom_idx.groupby(pc_key, as_index=False).first().set_index(pc_key, drop=False)

    # (CHANGE 1) Build UZIO Employment Status map (mandatory in UZIO)
    uz_emp_status_col = find_col(uzio_idx.columns, "Employment Status")
    if uz_emp_status_col is None:
        # Still allow tool to run, but column will be blank
        uz_emp_status_map = {}
    else:
        uz_emp_status_map = (
            uzio_idx[uz_emp_status_col]
            .astype(object)
            .where(~uzio_idx[uz_emp_status_col].isna(), "")
            .map(lambda v: "" if norm_blank(v) == "" else str(v))
            .to_dict()
        )

    # Mapping sanity: columns must exist
    def _col_exists(df_cols, col_name):
        return norm_colname(col_name).casefold() in {norm_colname(c).casefold() for c in df_cols}

    rows = []
    all_emps = sorted(set(uzio_idx.index.tolist()) | set(paycom_idx.index.tolist()))

    for emp in all_emps:
        emp = "" if norm_blank(emp) == "" else str(emp).strip()

        # Pay type for skip rules (kept as earlier requirement)
        pay_type_norm = get_pay_type_for_employee(emp, uzio_idx, paycom_idx, uz_key, pc_key)

        for _, m in mapping.iterrows():
            uz_col = m["UZIO_Column"]
            pc_col = m["PAYCOM_Column"]
            field = uz_col

            uz_exists = _col_exists(uzio_idx.columns, uz_col)
            pc_exists = _col_exists(paycom_idx.columns, pc_col)

            uz_val = ""
            pc_val = ""
            if emp in uzio_idx.index and uz_exists:
                uz_val = uzio_idx.loc[emp, uz_col]
            if emp in paycom_idx.index and pc_exists:
                pc_val = paycom_idx.loc[emp, pc_col]

            # Determine status (keep same structure)
            if emp not in paycom_idx.index and emp in uzio_idx.index:
                status = "MISSING_IN_PAYCOM"
            elif emp in paycom_idx.index and emp not in uzio_idx.index:
                status = "MISSING_IN_UZIO"
            elif not pc_exists:
                status = "PAYCOM_COLUMN_MISSING"
            elif not uz_exists:
                status = "UZIO_COLUMN_MISSING"
            else:
                # Skip rule: salaried -> don't compare hourly rate
                if should_skip_field_for_employee(field, pay_type_norm):
                    status = "OK"
                else:
                    # Special-case: Termination Reason rules
                    f_norm = norm_colname(field).casefold()
                    if any(k in f_norm for k in TERMINATION_REASON_KEYWORDS):
                        if termination_reason_ok(uz_val, pc_val):
                            status = "OK"
                        else:
                            uz_n = norm_value(uz_val, field)
                            pc_n = norm_value(pc_val, field)
                            if (uz_n == pc_n) or (uz_n == "" and pc_n == ""):
                                status = "OK"
                            elif uz_n == "" and pc_n != "":
                                status = "UZIO_MISSING_VALUE"
                            elif uz_n != "" and pc_n == "":
                                status = "PAYCOM_MISSING_VALUE"
                            else:
                                status = "MISMATCH"
                    else:
                        uz_n = norm_value(uz_val, field)
                        pc_n = norm_value(pc_val, field)
                        if (uz_n == pc_n) or (uz_n == "" and pc_n == ""):
                            status = "OK"
                        elif uz_n == "" and pc_n != "":
                            status = "UZIO_MISSING_VALUE"
                        elif uz_n != "" and pc_n == "":
                            status = "PAYCOM_MISSING_VALUE"
                        else:
                            status = "MISMATCH"

            # (CHANGE 1) Add Employment Status column after Field (always from UZIO)
            rows.append(
                {
                    "Employee": emp,
                    "Field": field,
                    "Employment Status": uz_emp_status_map.get(emp, ""),  # context column
                    "UZIO_Value": uz_val,
                    "PAYCOM_Value": pc_val,
                    "PAYCOM_SourceOfTruth_Status": status,
                }
            )

    comparison_detail = pd.DataFrame(
        rows,
        columns=[
            "Employee",
            "Field",
            "Employment Status",  # (CHANGE 1) column position enforced here
            "UZIO_Value",
            "PAYCOM_Value",
            "PAYCOM_SourceOfTruth_Status",
        ],
    )

    # Field summary
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
                values="Employee",
                aggfunc="count",
                fill_value=0,
            )
            .reindex(columns=statuses, fill_value=0)
            .reset_index()
        )
        field_summary_by_status["Total"] = field_summary_by_status[statuses].sum(axis=1)
    else:
        field_summary_by_status = pd.DataFrame(columns=["Field"] + statuses + ["Total"])

    # Summary
    uzio_emp = set(uzio_idx.index.tolist())
    paycom_emp = set(paycom_idx.index.tolist())
    fields_compared = int(mapping.shape[0])

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total UZIO Employees",
                "Total PAYCOM Employees",
                "Employees in both",
                "Employees only in UZIO",
                "Employees only in PAYCOM",
                "Total UZIO Records",
                "Total PAYCOM Records",
                "Fields Compared",
                "Total Comparisons (field-level rows)",
            ],
            "Value": [
                len(uzio_emp),
                len(paycom_emp),
                len(uzio_emp & paycom_emp),
                len(uzio_emp - paycom_emp),
                len(paycom_emp - uzio_emp),
                len(uzio_idx),
                len(paycom_idx),
                fields_compared,
                len(comparison_detail),
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
            report_bytes = build_report(uploaded_file.getvalue())

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
