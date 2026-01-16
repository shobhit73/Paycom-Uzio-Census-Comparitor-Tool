import io
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation

import numpy as np
import pandas as pd
import streamlit as st


# =========================
# CONFIG
# =========================
APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET_NAME = "Uzio Data"
PAYCOM_SHEET_NAME = "Paycom Data"
MAPPING_SHEET_NAME = "Mapping Sheet"

DETAIL_SHEET = "Comparison_Detail_AllFields"
SUMMARY_SHEET = "Summary"
FIELD_SUMMARY_SHEET = "Field_Summary_By_Status"


# =========================
# UI
# =========================
st.set_page_config(page_title=APP_TITLE, layout="centered")
st.title(APP_TITLE)
st.write("Upload the Excel workbook (.xlsx) with 3 tabs: Uzio Data, Paycom Data, and Mapping Sheet.")


# =========================
# HELPERS
# =========================
def norm_colname(c) -> str:
    if c is None:
        return ""
    s = str(c).replace("\n", " ").replace("\r", " ").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_blank(x) -> bool:
    if x is None:
        return True
    if isinstance(x, float) and np.isnan(x):
        return True
    s = str(x).strip()
    return s == "" or s.lower() in {"nan", "null", "none"}


def to_str(x) -> str:
    if is_blank(x):
        return ""
    return str(x).strip()


def norm_key_series(s: pd.Series) -> pd.Series:
    # Handles excel numbers like 123.0
    return s.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()


def digits_only(x: str) -> str:
    return re.sub(r"\D", "", to_str(x))


def norm_phone(x) -> str:
    d = digits_only(x)
    # Paycom sometimes includes country code "1"
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    # compare last 10 digits
    return d[-10:] if len(d) >= 10 else d


def norm_zip(x) -> str:
    d = digits_only(x)
    return d[:5]


def try_decimal(x):
    s = to_str(x)
    if s == "":
        return None
    s = s.replace(",", "")
    try:
        d = Decimal(s)
        # normalize to remove trailing zeros (150000.00 == 150000)
        return d.normalize()
    except (InvalidOperation, ValueError):
        return None


def num_equal(a, b) -> bool:
    da = try_decimal(a)
    db = try_decimal(b)
    if da is None or db is None:
        return False
    return da == db


def norm_simple_text(x: str) -> str:
    # normalize for comparisons like Full-Time vs Full Time, Jr. vs JR
    s = to_str(x).casefold()
    s = s.replace(".", "")
    s = re.sub(r"[\s\-]+", "", s)  # remove spaces + hyphens
    return s


def contains_word(haystack: str, needle: str) -> bool:
    return needle.casefold() in to_str(haystack).casefold()


def find_sheet_case_insensitive(xls: pd.ExcelFile, expected: str) -> str:
    # Allow exact match OR case-insensitive match
    for s in xls.sheet_names:
        if s == expected:
            return s
    for s in xls.sheet_names:
        if s.casefold() == expected.casefold():
            return s
    raise ValueError(f"Sheet not found: '{expected}'. Available: {xls.sheet_names}")


def choose_key_column(cols, preferred_list):
    cols_cf = {c.casefold(): c for c in cols}
    for p in preferred_list:
        if p.casefold() in cols_cf:
            return cols_cf[p.casefold()]
    # fallback: any column containing "employee"
    for c in cols:
        if "employee" in c.casefold():
            return c
    return None


# =========================
# MAPPING SHEET READERS
# =========================
def read_column_mapping(mapping_df: pd.DataFrame, uz_cols, pc_cols):
    """
    Expects two columns in mapping sheet for field mapping:
      - UZIO Column
      - Paycom Column
    (case-insensitive)
    """
    m = mapping_df.copy()
    m.columns = [norm_colname(c) for c in m.columns]

    # detect mapping header names
    uz_header = None
    pc_header = None
    for c in m.columns:
        if c.casefold() == "uzio column":
            uz_header = c
        if c.casefold() == "paycom column":
            pc_header = c
    if uz_header is None or pc_header is None:
        raise ValueError("Mapping Sheet must contain columns: 'UZIO Column' and 'Paycom Column'.")

    m[uz_header] = m[uz_header].map(norm_colname)
    m[pc_header] = m[pc_header].map(norm_colname)

    # Remove key mappings (employee id etc.) from column mapping
    m["_uz_norm"] = m[uz_header].str.casefold()
    m = m[~m["_uz_norm"].isin(["employee", "employee id", "employee_code", "employee code"])].copy()

    # resolve paycom column names to exact header in paycom sheet
    pc_map = {c.casefold(): c for c in pc_cols}
    m["PAYCOM_Resolved"] = m[pc_header].apply(lambda x: pc_map.get(to_str(x).casefold(), ""))

    return m.rename(columns={uz_header: "UZIO Column", pc_header: "Paycom Column"})[["UZIO Column", "Paycom Column", "PAYCOM_Resolved"]]


def read_value_mapping_table(mapping_df: pd.DataFrame, left_col_name: str, right_col_name: str):
    """
    Reads a 2-column mapping like:
      Paycom Termination Reason | Uzio Termination Reason
    Returns dict: normalized_paycom_value -> uzio_value (original)
    """
    cols_cf = {c.casefold(): c for c in mapping_df.columns}
    lc = cols_cf.get(left_col_name.casefold())
    rc = cols_cf.get(right_col_name.casefold())
    if lc is None or rc is None:
        return {}

    sub = mapping_df[[lc, rc]].copy()
    sub = sub.dropna(how="all")
    out = {}
    for _, row in sub.iterrows():
        k = to_str(row[lc])
        v = to_str(row[rc])
        if k != "" and v != "":
            out[k.casefold()] = v
    return out


# =========================
# CORE COMPARISON RULES
# =========================
def compare_values(field, uz_val, pc_val, emp_pay_type, value_maps):
    """
    Returns one of:
      OK, MISMATCH, UZIO_MISSING_VALUE, PAYCOM_MISSING_VALUE, UZIO_COLUMN_MISSING, PAYCOM_COLUMN_MISSING
    """
    f = to_str(field)

    # 1) Original DOH should never be considered mismatch
    if f.casefold() == "original doh":
        return "OK"

    # Standard missing checks
    uz_blank = is_blank(uz_val)
    pc_blank = is_blank(pc_val)

    # If both blank -> OK
    if uz_blank and pc_blank:
        return "OK"

    # =========================
    # FIELD-SPECIFIC RULES
    # =========================
    fname = f.casefold()

    # Phone Number(Digits): ignore country code 1
    if "phone number" in fname:
        if uz_blank and not pc_blank:
            # still mismatch? user wants ignore country code only, not missing.
            # We'll treat missing as missing.
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        return "OK" if norm_phone(uz_val) == norm_phone(pc_val) else "MISMATCH"

    # Zipcode: first 5 digits must match
    if "zipcode" in fname:
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        return "OK" if norm_zip(uz_val) == norm_zip(pc_val) else "MISMATCH"

    # Employee Middle Initial: UZIO is initial, Paycom may be full middle name
    if "employee middle initial" in fname:
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        uz_first = to_str(uz_val)[:1].upper()
        pc_first = to_str(pc_val)[:1].upper()
        return "OK" if uz_first != "" and uz_first == pc_first else "MISMATCH"

    # Employee Suffix: Jr. == JR, etc
    if "employee suffix" in fname:
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        return "OK" if norm_simple_text(uz_val) == norm_simple_text(pc_val) else "MISMATCH"

    # Employment Type: Full Time == Full-Time
    if "employment type" in fname:
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        return "OK" if norm_simple_text(uz_val) == norm_simple_text(pc_val) else "MISMATCH"

    # Employment Status: Paycom "On Leave" == UZIO "Active"
    if fname == "employment status":
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        pc_norm = to_str(pc_val).casefold()
        uz_norm = to_str(uz_val).casefold()
        if pc_norm == "on leave" and uz_norm == "active":
            return "OK"
        return "OK" if uz_norm == pc_norm else "MISMATCH"

    # Termination Reason:
    # If both contain "voluntary" -> OK
    # If both contain "involuntary" -> OK
    # Else use mapping table (if exists)
    if fname == "termination reason":
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"

        uz_s = to_str(uz_val)
        pc_s = to_str(pc_val)

        if contains_word(uz_s, "voluntary") and contains_word(pc_s, "voluntary"):
            return "OK"
        if contains_word(uz_s, "involuntary") and contains_word(pc_s, "involuntary"):
            return "OK"

        term_map = value_maps.get("termination_reason_map", {})
        mapped = term_map.get(pc_s.casefold(), "")
        if mapped:
            return "OK" if norm_simple_text(uz_s) == norm_simple_text(mapped) else "MISMATCH"

        # fallback direct compare
        return "OK" if norm_simple_text(uz_s) == norm_simple_text(pc_s) else "MISMATCH"

    # Pay Type: Salaried == Salary
    if fname == "pay type":
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        uz_norm = norm_simple_text(uz_val)
        pc_norm = norm_simple_text(pc_val)
        # map salary <-> salaried
        if uz_norm == "salaried" and pc_norm == "salary":
            return "OK"
        if uz_norm == "salary" and pc_norm == "salaried":
            return "OK"
        return "OK" if uz_norm == pc_norm else "MISMATCH"

    # If person is salaried then don't compare hourly rate + working hours/week
    if "hourly pay rate" in fname or "working hours per week" in fname:
        if contains_word(emp_pay_type, "salary") or contains_word(emp_pay_type, "salaried"):
            return "OK"
        # hourly employee -> normal compare (numeric)
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        # numeric compare (ignore decimals)
        if num_equal(uz_val, pc_val):
            return "OK"
        # also allow text normalized compare
        return "OK" if norm_simple_text(uz_val) == norm_simple_text(pc_val) else "MISMATCH"

    # Annual Salary (Digits):
    # - If employee is hourly and UZIO blank => OK (not UZIO_MISSING_VALUE)
    if "annual salary" in fname:
        if contains_word(emp_pay_type, "hourly"):
            # UZIO should be blank; treat blank as OK even if Paycom has value
            return "OK" if uz_blank else ("OK" if num_equal(uz_val, pc_val) else "MISMATCH")
        # salaried -> compare numeric
        if uz_blank and not pc_blank:
            return "UZIO_MISSING_VALUE"
        if pc_blank and not uz_blank:
            return "PAYCOM_MISSING_VALUE"
        return "OK" if num_equal(uz_val, pc_val) else "MISMATCH"

    # Generic numeric equality: ignore decimals formatting
    if num_equal(uz_val, pc_val):
        return "OK"

    # Generic string compare (case-insensitive, trim)
    if norm_simple_text(uz_val) == norm_simple_text(pc_val):
        return "OK"

    # Missing (generic) if one side blank
    if uz_blank and not pc_blank:
        return "UZIO_MISSING_VALUE"
    if pc_blank and not uz_blank:
        return "PAYCOM_MISSING_VALUE"

    return "MISMATCH"


# =========================
# MAIN RUN
# =========================
def run_audit(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uz_sheet = find_sheet_case_insensitive(xls, UZIO_SHEET_NAME)
    pc_sheet = find_sheet_case_insensitive(xls, PAYCOM_SHEET_NAME)
    mp_sheet = find_sheet_case_insensitive(xls, MAPPING_SHEET_NAME)

    uzio = pd.read_excel(xls, sheet_name=uz_sheet, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=pc_sheet, dtype=object)
    mapping_df = pd.read_excel(xls, sheet_name=mp_sheet, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    mapping_df.columns = [norm_colname(c) for c in mapping_df.columns]

    # Key columns
    uz_key = choose_key_column(uzio.columns, ["Employee", "Employee ID", "Employee_Code", "Employee Code"])
    pc_key = choose_key_column(paycom.columns, ["Employee_Code", "Employee Code", "Employee", "Employee ID"])

    if uz_key is None:
        raise ValueError("UZIO key column not found (expected Employee / Employee ID / Employee_Code).")
    if pc_key is None:
        raise ValueError("Paycom key column not found (expected Employee_Code / Employee / Employee ID).")

    uzio[uz_key] = norm_key_series(uzio[uz_key])
    paycom[pc_key] = norm_key_series(paycom[pc_key])

    # Deduplicate keys by taking first record
    uzio = uzio.drop_duplicates(subset=[uz_key], keep="first")
    paycom = paycom.drop_duplicates(subset=[pc_key], keep="first")

    uz_by = uzio.set_index(uz_key, drop=False)
    pc_by = paycom.set_index(pc_key, drop=False)

    # Column mapping
    col_map = read_column_mapping(mapping_df, uzio.columns, paycom.columns)

    # Optional value mappings
    value_maps = {
        "termination_reason_map": read_value_mapping_table(mapping_df, "Paycom Termination Reason", "Uzio Termination Reason")
    }

    employees = sorted(set(uz_by.index) | set(pc_by.index))

    out_rows = []

    for emp in employees:
        uz_row = uz_by.loc[emp] if emp in uz_by.index else None
        pc_row = pc_by.loc[emp] if emp in pc_by.index else None

        # If multiple rows somehow, take first
        if isinstance(uz_row, pd.DataFrame):
            uz_row = uz_row.iloc[0]
        if isinstance(pc_row, pd.DataFrame):
            pc_row = pc_row.iloc[0]

        emp_status = to_str(uz_row.get("Employment Status", "")) if uz_row is not None else ""
        emp_pay_type = to_str(uz_row.get("Pay Type", "")) if uz_row is not None else ""

        for _, m in col_map.iterrows():
            uz_field = to_str(m["UZIO Column"])
            pc_col = to_str(m["PAYCOM_Resolved"])

            # Column missing checks
            if uz_field not in uzio.columns:
                status = "UZIO_COLUMN_MISSING"
                uz_val = ""
                pc_val = ""
            elif pc_col == "" or pc_col not in paycom.columns:
                status = "PAYCOM_COLUMN_MISSING"
                uz_val = to_str(uz_row.get(uz_field, "")) if uz_row is not None else ""
                pc_val = ""
            else:
                uz_val = to_str(uz_row.get(uz_field, "")) if uz_row is not None else ""
                pc_val = to_str(pc_row.get(pc_col, "")) if pc_row is not None else ""
                status = compare_values(uz_field, uz_val, pc_val, emp_pay_type, value_maps)

            out_rows.append({
                "Employee": emp,
                "Field": uz_field,
                "Employment Status": emp_status,
                "UZIO_Value": uz_val,
                "PAYCOM_Value": pc_val,
                "PAYCOM_SourceOfTruth_Status": status
            })

    detail_df = pd.DataFrame(out_rows)

    # Summary
    status_counts = detail_df["PAYCOM_SourceOfTruth_Status"].value_counts(dropna=False).to_dict()
    summary_rows = [{"Metric": k, "Value": v} for k, v in status_counts.items()]
    summary_rows.insert(0, {"Metric": "Total Rows Compared", "Value": len(detail_df)})
    summary_df = pd.DataFrame(summary_rows)

    # Field Summary by Status
    field_summary_df = (
        detail_df.groupby(["Field", "PAYCOM_SourceOfTruth_Status"])
        .size()
        .reset_index(name="Count")
        .sort_values(["Field", "PAYCOM_SourceOfTruth_Status"])
    )

    # Write output excel
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, SUMMARY_SHEET, index=False)
        field_summary_df.to_excel(writer, FIELD_SUMMARY_SHEET, index=False)
        detail_df.to_excel(writer, DETAIL_SHEET, index=False)

    return out.getvalue()


# =========================
# STREAMLIT APP
# =========================
uploaded = st.file_uploader("Upload Excel workbook", type=["xlsx"])

if st.button("Run Audit", disabled=(uploaded is None)):
    try:
        report_bytes = run_audit(uploaded.getvalue())
        filename = f"UZIO_vs_PAYCOM_Comparison_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.success("Audit completed.")
        st.download_button(
            "Download Report",
            data=report_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Failed: {e}")
