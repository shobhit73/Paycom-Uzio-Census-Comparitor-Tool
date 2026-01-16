# app.py
import io
import re
from datetime import datetime, date
from decimal import Decimal, InvalidOperation

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom vs UZIO – Census Audit Tool
# INPUT workbook tabs:
#   - Uzio Data
#   - Paycom Data
#   - Mapping Sheet
#
# OUTPUT tabs:
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

def resolve_col_label(label: str, cols_all) -> str:
    """
    Mapping sheet might have slightly different labels, commas, (or ...).
    Resolve to actual column name present in the sheet.
    """
    if label is None:
        return ""
    raw = str(label).strip()
    raw = raw.replace("’", "'").replace("“", '"').replace("”", '"')
    raw = raw.strip().strip(",")
    if raw == "":
        return ""

    cols_norm = {norm_colname(c).casefold(): c for c in cols_all}

    direct = norm_colname(raw).casefold()
    if direct in cols_norm:
        return cols_norm[direct]

    parts = re.split(r"\(|\)|\bor\b|/|,|;", raw, flags=re.IGNORECASE)
    parts = [norm_colname(p) for p in parts if norm_colname(p)]
    extra = []
    for p in parts:
        extra.extend([norm_colname(x) for x in re.split(r"\s[-–]\s", p) if norm_colname(x)])
    parts = parts + extra

    for p in parts:
        k = norm_colname(p).casefold()
        if k in cols_norm:
            return cols_norm[k]

    # fallback contains
    for k_norm, actual in cols_norm.items():
        if k_norm and (k_norm in direct or direct in k_norm):
            return actual

    return ""

def digits_only(x):
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

def norm_phone_digits(x):
    """
    Normalize phone to digits only.
    - Drop leading US country code '1' if present (11 digits).
    - If longer than 10, keep last 10.
    """
    d = digits_only(x)
    if d == "":
        return ""
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) > 10:
        d = d[-10:]
    return d

def norm_zip_first5(x):
    d = digits_only(x)
    if d == "":
        return ""
    if 0 < len(d) < 5:
        d = d.zfill(5)
    return d[:5]

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

def safe_decimal(x):
    x = norm_blank(x)
    if x == "":
        return None
    if isinstance(x, (int, np.integer)):
        return Decimal(int(x))
    if isinstance(x, (float, np.floating)):
        # avoid float artifacts by string conversion
        s = str(x)
        try:
            return Decimal(s)
        except Exception:
            return None
    s = str(x).strip().replace(",", "").replace("$", "")
    # strip trailing .0 patterns
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return None

def approx_equal_decimal(a: Decimal, b: Decimal) -> bool:
    # Ignore decimal formatting differences: 80 vs 80.0, 150000 vs 150000.00
    if a is None or b is None:
        return False
    return a == b

def norm_spaces_punct(s: str) -> str:
    s = norm_blank(s)
    if s == "":
        return ""
    s = str(s)
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[-_/]", " ", s)           # treat hyphens etc as spaces
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)  # drop punctuation
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def norm_middle_initial(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).strip()
    # find first alphabet char
    m = re.search(r"[A-Za-z]", s)
    return (m.group(0).upper() if m else "").strip()

def norm_suffix(x):
    # "Jr." == "JR" == "jr"
    return norm_spaces_punct(x)

def norm_pay_type(x):
    s = norm_spaces_punct(x)
    if s == "":
        return ""
    # normalize salaried/salary
    if "salar" in s:
        return "salary"
    if "hour" in s:
        return "hourly"
    return s

def norm_employment_status(x):
    s = norm_spaces_punct(x)
    if s == "":
        return ""
    # Paycom: "On Leave" should be treated as ACTIVE in UZIO comparisons
    if s in {"on leave", "leave"}:
        return "active"
    return s

def norm_termination_reason(x):
    return norm_spaces_punct(x)

def termination_reason_equivalent(uzio_val, paycom_val) -> bool:
    u = norm_termination_reason(uzio_val)
    p = norm_termination_reason(paycom_val)
    if u == "" and p == "":
        return True
    # keyword rule requested: if both contain voluntary OR both contain involuntary -> OK
    u_has_vol = "voluntary" in u
    p_has_vol = "voluntary" in p
    u_has_invol = "involuntary" in u
    p_has_invol = "involuntary" in p
    if u_has_invol and p_has_invol:
        return True
    if u_has_vol and p_has_vol:
        return True
    return False

# Add DOH as date keyword so "Original DOH" is treated like a date
DATE_KEYWORDS = {"date", "dob", "birth", "effective", "doh", "hire"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
PHONE_KEYWORDS = {"phone", "mobile"}
NUMERIC_HINTS = {"salary", "rate", "hours", "amount", "percent", "percentage", "digits"}

def norm_value_for_field(x, field_name: str):
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    # Phone / Zip
    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)
    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    # Dates (including DOH)
    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    # Middle Initial
    if "middle" in f and "initial" in f:
        return norm_middle_initial(x)

    # Suffix
    if "suffix" in f:
        return norm_suffix(x)

    # Employment Type (Full time vs Full-Time)
    if "employment type" in f:
        return norm_spaces_punct(x)

    # Pay Type mapping
    if f.strip() == "pay type" or "pay type" in f:
        return norm_pay_type(x)

    # Employment Status mapping (On Leave -> Active)
    if "employment status" in f:
        return norm_employment_status(x)

    # Numeric-ish (ignore decimal display differences)
    if any(k in f for k in NUMERIC_HINTS):
        d = safe_decimal(x)
        if d is not None:
            return d
        return norm_spaces_punct(x)

    # default string normalization
    return norm_spaces_punct(x)

# ---------- Mapping sheet ----------
def read_mapping_sheet(xls: pd.ExcelFile, sheet_name: str, uzio_cols: list, paycom_cols: list) -> pd.DataFrame:
    m = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
    m.columns = [norm_colname(c) for c in m.columns]

    uz_col_name = None
    pc_col_name = None
    for c in m.columns:
        c_norm = norm_colname(c).casefold()
        if c_norm in {"uzio coloumn", "uzio column"}:
            uz_col_name = c
        if c_norm in {"paycom coloumn", "paycom column"}:
            pc_col_name = c

    if uz_col_name is None or pc_col_name is None:
        raise ValueError("Mapping Sheet must contain columns: 'UZIO Column' and 'Paycom Column'.")

    m[uz_col_name] = m[uz_col_name].map(norm_colname)
    m[pc_col_name] = m[pc_col_name].map(norm_colname)

    m = m.dropna(subset=[uz_col_name, pc_col_name]).copy()
    m = m[(m[uz_col_name] != "") & (m[pc_col_name] != "")]
    m = m.drop_duplicates(subset=[uz_col_name], keep="first").copy()

    m["UZIO_Column"] = m[uz_col_name]
    m["PAYCOM_Label"] = m[pc_col_name]
    m["PAYCOM_Resolved_Column"] = m["PAYCOM_Label"].map(lambda x: resolve_col_label(x, paycom_cols))

    # exclude key row from comparisons (it is only key)
    m["_uz_norm"] = m["UZIO_Column"].map(lambda x: norm_colname(x).casefold())
    m = m[m["_uz_norm"] not in ["employee id", "employee", "employee_code", "employee code"]].copy()
    m.drop(columns=["_uz_norm"], inplace=True, errors="ignore")

    return m

# ---------- Core comparison ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    # Read sheets
    uzio = pd.read_excel(xls, sheet_name=UZIO_SHEET, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=PAYCOM_SHEET, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=MAP_SHEET, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    # Keys
    UZIO_KEY = find_col(uzio.columns, "Employee ID", "EmployeeID", "Employee", "Employee Code", "Employee_Code")
    if UZIO_KEY is None:
        raise ValueError("UZIO key column not found (expected 'Employee ID' / 'Employee').")

    PAYCOM_KEY = find_col(paycom.columns, "Employee_Code", "Employee Code", "Employee ID", "Employee", "EmployeeID")
    if PAYCOM_KEY is None:
        raise ValueError("Paycom key column not found (expected 'Employee_Code' / 'Employee ID' / 'Employee').")

    # Normalize keys
    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    # Re-read mapping using common helper (so it can resolve columns)
    # We'll pass the already loaded mapping via xls to preserve the same logic
    # but easiest is to build a temp ExcelFile-compatible object: use xls again.
    map_df = pd.read_excel(xls, sheet_name=MAP_SHEET, dtype=object)
    map_df.columns = [norm_colname(c) for c in map_df.columns]

    # Build mapping table in the same format
    # Create a faux object so we can use the same resolve logic
    tmp = io.BytesIO(file_bytes)
    xls2 = pd.ExcelFile(tmp, engine="openpyxl")
    map_tbl = read_mapping_sheet(
        xls2,
        MAP_SHEET,
        uzio_cols=list(uzio.columns),
        paycom_cols=list(paycom.columns),
    )

    # Only compare rows where paycom column was resolved
    map_tbl_valid = map_tbl.copy()

    # Build row lookup by key
    uzio_by = uzio.set_index(UZIO_KEY, drop=False)
    paycom_by = paycom.set_index(PAYCOM_KEY, drop=False)

    all_emp = sorted(set(uzio_by.index.astype(str)).union(set(paycom_by.index.astype(str))) - {""})

    # Determine columns for context
    uzio_emp_status_col = find_col(uzio.columns, "Employment Status")
    uzio_pay_type_col = find_col(uzio.columns, "Pay Type")

    rows = []
    for emp in all_emp:
        uz_row = uzio_by.loc[emp] if emp in uzio_by.index else None
        pc_row = paycom_by.loc[emp] if emp in paycom_by.index else None

        # If duplicate keys, pick first row (keep tool stable)
        if isinstance(uz_row, pd.DataFrame):
            uz_row = uz_row.iloc[0]
        if isinstance(pc_row, pd.DataFrame):
            pc_row = pc_row.iloc[0]

        # Context: Employment Status column (after Field)
        emp_status_ctx = ""
        if uz_row is not None and uzio_emp_status_col in uzio.columns:
            emp_status_ctx = uz_row.get(uzio_emp_status_col, "")
        elif pc_row is not None:
            # fallback: if paycom has similar column name
            pc_status_col = find_col(paycom.columns, "Employment Status", "EmploymentStatus")
            if pc_status_col:
                emp_status_ctx = pc_row.get(pc_status_col, "")
        emp_status_ctx = str(norm_blank(emp_status_ctx) or "").strip()

        # Context: Pay Type (used for conditional comparisons)
        pay_type_ctx_raw = ""
        if uz_row is not None and uzio_pay_type_col in uzio.columns:
            pay_type_ctx_raw = uz_row.get(uzio_pay_type_col, "")
        else:
            # fallback to paycom pay type if needed
            pc_pay_type_col = find_col(paycom.columns, "Pay Type")
            if pc_pay_type_col and pc_row is not None:
                pay_type_ctx_raw = pc_row.get(pc_pay_type_col, "")
        pay_type_ctx = norm_pay_type(pay_type_ctx_raw)

        for _, m in map_tbl_valid.iterrows():
            uz_field = m["UZIO_Column"]
            pc_col = m["PAYCOM_Resolved_Column"]

            # raw values
            uz_val = ""
            pc_val = ""

            if uz_row is not None and uz_field in uzio.columns:
                uz_val = uz_row.get(uz_field, "")
            if pc_row is not None and pc_col and pc_col in paycom.columns:
                pc_val = pc_row.get(pc_col, "")

            # statuses: missing record cases
            if emp not in uzio_by.index and emp in paycom_by.index:
                status = "MISSING_IN_UZIO"
            elif emp in uzio_by.index and emp not in paycom_by.index:
                status = "MISSING_IN_PAYCOM"
            elif pc_col == "" or (pc_col not in paycom.columns):
                status = "PAYCOM_COLUMN_MISSING"
            elif uz_field not in uzio.columns:
                status = "UZIO_COLUMN_MISSING"
            else:
                f_norm = norm_colname(uz_field).casefold()

                # -------- Special rules (requested) --------

                # Termination Reason: voluntary/involuntary keyword match -> OK
                if "termination reason" in f_norm:
                    if termination_reason_equivalent(uz_val, pc_val):
                        status = "OK"
                    else:
                        u_n = norm_value_for_field(uz_val, uz_field)
                        p_n = norm_value_for_field(pc_val, uz_field)
                        status = "OK" if u_n == p_n else "MISMATCH"

                else:
                    # Pay type mapping: Salaried == Salary
                    # (handled by norm_value_for_field)

                    # Conditional comparisons by Pay Type:
                    # - Salaried: ignore missing/mismatch for Hourly Pay Rate & Working Hours per Week when UZIO blank
                    if pay_type_ctx == "salary":
                        if ("hourly pay rate" in f_norm) or ("hourly rate" in f_norm) or ("working hours per week" in f_norm):
                            # if UZIO blank, treat OK regardless of paycom value (including 0)
                            if str(norm_blank(uz_val) or "").strip() == "":
                                status = "OK"
                            else:
                                u_n = norm_value_for_field(uz_val, uz_field)
                                p_n = norm_value_for_field(pc_val, uz_field)
                                status = "OK" if u_n == p_n else "MISMATCH"
                        else:
                            # normal compare
                            u_n = norm_value_for_field(uz_val, uz_field)
                            p_n = norm_value_for_field(pc_val, uz_field)

                            # handle numeric decimals robustly
                            if isinstance(u_n, Decimal) and isinstance(p_n, Decimal):
                                status = "OK" if approx_equal_decimal(u_n, p_n) else "MISMATCH"
                            else:
                                if (u_n == p_n) or (u_n == "" and p_n == ""):
                                    status = "OK"
                                elif u_n == "" and p_n != "":
                                    status = "UZIO_MISSING_VALUE"
                                elif u_n != "" and p_n == "":
                                    status = "PAYCOM_MISSING_VALUE"
                                else:
                                    status = "MISMATCH"

                    elif pay_type_ctx == "hourly":
                        # Hourly: if Annual Salary is blank in UZIO, that's OK
                        if "annual salary" in f_norm:
                            if str(norm_blank(uz_val) or "").strip() == "":
                                status = "OK"
                            else:
                                u_n = norm_value_for_field(uz_val, uz_field)
                                p_n = norm_value_for_field(pc_val, uz_field)
                                if isinstance(u_n, Decimal) and isinstance(p_n, Decimal):
                                    status = "OK" if approx_equal_decimal(u_n, p_n) else "MISMATCH"
                                else:
                                    status = "OK" if u_n == p_n else "MISMATCH"
                        else:
                            u_n = norm_value_for_field(uz_val, uz_field)
                            p_n = norm_value_for_field(pc_val, uz_field)
                            if isinstance(u_n, Decimal) and isinstance(p_n, Decimal):
                                status = "OK" if approx_equal_decimal(u_n, p_n) else "MISMATCH"
                            else:
                                if (u_n == p_n) or (u_n == "" and p_n == ""):
                                    status = "OK"
                                elif u_n == "" and p_n != "":
                                    status = "UZIO_MISSING_VALUE"
                                elif u_n != "" and p_n == "":
                                    status = "PAYCOM_MISSING_VALUE"
                                else:
                                    status = "MISMATCH"
                    else:
                        # Unknown pay type: normal compare
                        u_n = norm_value_for_field(uz_val, uz_field)
                        p_n = norm_value_for_field(pc_val, uz_field)
                        if isinstance(u_n, Decimal) and isinstance(p_n, Decimal):
                            status = "OK" if approx_equal_decimal(u_n, p_n) else "MISMATCH"
                        else:
                            if (u_n == p_n) or (u_n == "" and p_n == ""):
                                status = "OK"
                            elif u_n == "" and p_n != "":
                                status = "UZIO_MISSING_VALUE"
                            elif u_n != "" and p_n == "":
                                status = "PAYCOM_MISSING_VALUE"
                            else:
                                status = "MISMATCH"

            rows.append(
                {
                    "Employee": emp,
                    "Field": uz_field,
                    "Employment Status": emp_status_ctx,  # context column (after Field)
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
            "Employment Status",
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
    uzio_emp = set(uzio[UZIO_KEY].dropna().map(str)) if UZIO_KEY in uzio.columns else set()
    pc_emp = set(paycom[PAYCOM_KEY].dropna().map(str)) if PAYCOM_KEY in paycom.columns else set()

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total UZIO Employees",
                "Total Paycom Employees",
                "Employees in both",
                "Employees only in UZIO",
                "Employees only in Paycom",
                "Total UZIO Records",
                "Total Paycom Records",
                "Fields Compared",
                "Total Comparisons (field-level rows)",
            ],
            "Value": [
                len(uzio_emp),
                len(pc_emp),
                len(uzio_emp & pc_emp),
                len(uzio_emp - pc_emp),
                len(pc_emp - uzio_emp),
                len(uzio),
                len(paycom),
                int(map_tbl_valid.shape[0]),
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
        out_name = f"UZIO_vs_PAYCOM_Comparison_Report_PAYCOM_SourceOfTruth_{ts}.xlsx"

        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
