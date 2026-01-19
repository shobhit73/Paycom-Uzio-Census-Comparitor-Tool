# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom vs UZIO – Census Audit Tool
# INPUT workbook tabs (single file):
#   - Uzio Data
#   - Paycom Data
#   - Mapping Sheet   (UZIO Column -> Paycom Column)
#
# OUTPUT workbook tabs:
#   - Summary
#   - Field_Summary_By_Status
#   - Comparison_Detail_AllFields
#
# Key rules included:
#   ✅ Dates compare as DATE (ignore time part)
#   ✅ Termination Reason: voluntary/involuntary keyword logic + UZIO=Other => OK
#   ✅ Pay Type: UZIO "Salaried" == Paycom "Salary" => OK
#   ✅ Suffix: Jr. == JR => OK
#   ✅ Employment Type: Full Time == Full-Time => OK
#   ✅ Middle initial: UZIO first letter matches Paycom full middle name => OK
#   ✅ Employment Status: Paycom "On Leave" treated as "Active" => OK
#   ✅ Numerics: 150000.00 == 150000, 80.0 == 80 => OK (tolerance compare)
#
# Pay Type driven ignores (IMPORTANT):
#   ✅ If employee is HOURLY: ignore Annual Salary fields (UZIO blank is OK)
#   ✅ If employee is SALARIED: ignore Hourly Pay Rate AND Working Hours per Week
#      (UZIO leaves these blank => should be OK, not UZIO_MISSING_VALUE)
#
# Adds extra output column: "Employment Status" right after "Field"
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET_CANDIDATES = ["Uzio Data", "UZIO Data", "Uzio", "UZIO"]
PAYCOM_SHEET_CANDIDATES = ["Paycom Data", "PAYCOM Data", "Paycom", "PAYCOM"]
MAP_SHEET_CANDIDATES = ["Mapping Sheet", "Mapping", "Mapping_Sheet", "MappingSheet"]

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

def find_col(df_cols, *candidate_names):
    norm_map = {norm_colname(c).casefold(): c for c in df_cols}
    for cand in candidate_names:
        key = norm_colname(cand).casefold()
        if key in norm_map:
            return norm_map[key]
    return None

def norm_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

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

def as_float_or_none(x):
    x = norm_blank(x)
    if x == "":
        return None
    if isinstance(x, (int, float, np.integer, np.floating)):
        try:
            return float(x)
        except Exception:
            return None
    if isinstance(x, str):
        s = x.strip().replace(",", "").replace("$", "")
        if s == "":
            return None
        try:
            return float(s)
        except Exception:
            return None
    return None

def normalize_space_and_case(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).strip()
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def normalize_employment_type(x):
    s = normalize_space_and_case(x)
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_suffix(x):
    s = normalize_space_and_case(x)
    s = re.sub(r"[^a-z0-9]", "", s)  # remove punctuation/spaces
    return s

def normalize_phone(x):
    s = norm_blank(x)
    if s == "":
        return ""
    # remove all non-digits
    digits = re.sub(r"[^0-9]", "", str(s))
    # if 11 digits and starts with 1, remove leading 1
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits

def first_alpha_char(x):
    s = norm_blank(x)
    if s == "":
        return ""
    txt = str(s).strip()
    m = re.search(r"[A-Za-z]", txt)
    return m.group(0).casefold() if m else ""

def normalize_middle_initial(uzio_val, paycom_val):
    # UZIO has 'M', Paycom has 'MICHELLE' => OK if first letter matches
    u = first_alpha_char(uzio_val)
    p = first_alpha_char(paycom_val)
    return u != "" and p != "" and u == p

def canonical_pay_type(x):
    s = normalize_space_and_case(x)
    if s == "":
        return ""
    if "hour" in s:
        return "hourly"
    if "salar" in s or "salary" in s:
        return "salaried"
    return s

def canonical_employment_status(x):
    # Paycom "On Leave" treated as "Active"
    s = normalize_space_and_case(x)
    if s == "":
        return ""
    if "on leave" in s:
        return "active"
    if s in {"active", "activated"}:
        return "active"
    return s

def termination_reason_equal(uzio_val, paycom_val):
    uz = normalize_space_and_case(uzio_val)
    pc = normalize_space_and_case(paycom_val)

    if uz == "" and pc == "":
        return True

    # UZIO "Other" is acceptable for any Paycom reason
    if uz == "other":
        return True

    # If either has involuntary, both must have involuntary
    if ("involuntary" in uz) or ("involuntary" in pc):
        return ("involuntary" in uz) and ("involuntary" in pc)

    # If either has voluntary, both must have voluntary
    if ("voluntary" in uz) or ("voluntary" in pc):
        return ("voluntary" in uz) and ("voluntary" in pc)

    return uz == pc

def resolve_sheet_name(xls: pd.ExcelFile, candidates):
    existing_norm = {norm_colname(s).casefold(): s for s in xls.sheet_names}
    for c in candidates:
        k = norm_colname(c).casefold()
        if k in existing_norm:
            return existing_norm[k]
    return None

def resolve_paycom_col_label(label: str, paycom_cols_all) -> str:
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

    extra = []
    for p in parts:
        extra.extend([norm_colname(x) for x in re.split(r"\s[-–]\s", p) if norm_colname(x)])
    parts = parts + extra

    for p in parts:
        k = norm_colname(p).casefold()
        if k in pay_norm:
            return pay_norm[k]

    for k_norm, actual in pay_norm.items():
        if k_norm and (k_norm in direct or direct in k_norm):
            return actual

    return ""

def read_mapping_sheet(xls: pd.ExcelFile, sheet_name: str, paycom_cols_all: list) -> pd.DataFrame:
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
    m["PAYCOM_Label"] = m[pc_col_name]
    m["PAYCOM_Resolved_Column"] = m["PAYCOM_Label"].map(lambda x: resolve_paycom_col_label(x, paycom_cols_all))

    # exclude Employee ID/Employee Code from comparisons (key only)
    m["_uz_norm"] = m["UZIO_Column"].map(lambda x: norm_colname(x).casefold())
    m = m[~m["_uz_norm"].isin({"employee id", "employee", "employee_code", "employee code"})].copy()
    m.drop(columns=["_uz_norm"], inplace=True)

    return m

def should_ignore_field_for_paytype(field_name: str, pay_type_canon: str) -> bool:
    """
    Pay-type based ignore rules (as per your requirement):
      - HOURLY employees: ignore annual salary fields
      - SALARIED employees: ignore hourly pay rate AND working hours per week
    """
    f = norm_colname(field_name).casefold()
    pt = (pay_type_canon or "").casefold()

    if pt == "hourly":
        if "annual salary" in f:
            return True

    if pt == "salaried":
        # ignore Hourly Pay Rate (covers: "Hourly Pay Rate", "Hourly Rate", etc.)
        if ("hourly" in f and "rate" in f):
            return True

        # ignore Working Hours per Week (covers: "Working Hours per Week(Digits)", "Hours per Week", etc.)
        if ("hours per week" in f) or ("working hours" in f):
            return True

    return False

def normalized_compare(field_name: str, uzio_val, paycom_val) -> bool:
    f = norm_colname(field_name).casefold()

    if "termination reason" in f:
        return termination_reason_equal(uzio_val, paycom_val)

    if "employment status" in f:
        return canonical_employment_status(uzio_val) == canonical_employment_status(paycom_val)

    if "pay type" in f:
        return canonical_pay_type(uzio_val) == canonical_pay_type(paycom_val)

    if "employment type" in f:
        return normalize_employment_type(uzio_val) == normalize_employment_type(paycom_val)

    if ("middle" in f) and ("initial" in f):
        if normalize_middle_initial(uzio_val, paycom_val):
            return True
        return first_alpha_char(uzio_val) == first_alpha_char(paycom_val)

    if "suffix" in f:
        return normalize_suffix(uzio_val) == normalize_suffix(paycom_val)

    if "ssn" in f:
        # Normalize SSN: digits only, remove leading zeros
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p

    if "phone" in f:
        # Normalize Phone: digits only, remove leading zeros
        u = normalize_phone(uzio_val).lstrip("0")
        p = normalize_phone(paycom_val).lstrip("0")
        return u == p

    if "zip" in f:
        # Normalize Zip: digits only (simple), remove leading zeros
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p

    # Date-ish fields (including DOH)
    if any(k in f for k in ["date", "dob", "birth", "effective", "doh", "hire", "termination"]):
        return try_parse_date(uzio_val) == try_parse_date(paycom_val)

    # Numeric-ish fields
    if any(k in f for k in ["salary", "rate", "hours", "amount", "percent", "percentage", "digits"]):
        fa = as_float_or_none(uzio_val)
        fb = as_float_or_none(paycom_val)
        if fa is not None and fb is not None:
            return abs(fa - fb) <= 1e-9
        return normalize_space_and_case(uzio_val) == normalize_space_and_case(paycom_val)

    return normalize_space_and_case(uzio_val) == normalize_space_and_case(paycom_val)

# ---------- Core comparison ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio_sheet = resolve_sheet_name(xls, UZIO_SHEET_CANDIDATES)
    paycom_sheet = resolve_sheet_name(xls, PAYCOM_SHEET_CANDIDATES)
    map_sheet = resolve_sheet_name(xls, MAP_SHEET_CANDIDATES)

    if uzio_sheet is None:
        raise ValueError("UZIO sheet not found. Expected a tab like 'Uzio Data'.")
    if paycom_sheet is None:
        raise ValueError("Paycom sheet not found. Expected a tab like 'Paycom Data'.")
    if map_sheet is None:
        raise ValueError("Mapping sheet not found. Expected a tab like 'Mapping Sheet' or 'Mapping'.")

    uzio = pd.read_excel(xls, sheet_name=uzio_sheet, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=paycom_sheet, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]

    # keys (robust)
    UZIO_KEY = find_col(
        uzio.columns,
        "Employee ID", "EmployeeID", "Employee Id", "Employee",
        "Employee_Code", "Employee Code"
    )
    if UZIO_KEY is None:
        raise ValueError("UZIO key column not found (expected 'Employee ID'/'Employee'/'Employee_Code').")

    PAYCOM_KEY = find_col(
        paycom.columns,
        "Employee_Code", "Employee Code",
        "Employee ID", "EmployeeID", "Employee Id", "Employee"
    )
    if PAYCOM_KEY is None:
        raise ValueError("Paycom key column not found (expected 'Employee_Code'/'Employee ID'/'Employee').")

    # normalize keys
    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    # mapping sheet
    mapping = read_mapping_sheet(xls, map_sheet, list(paycom.columns))
    mapping = mapping[mapping["PAYCOM_Resolved_Column"] != ""].copy()

    # employment status context map (prefer UZIO)
    uzio_emp_status_col = find_col(uzio.columns, "Employment Status")
    paycom_emp_status_col = find_col(paycom.columns, "Employment Status")

    uzio_status_map = {}
    if uzio_emp_status_col is not None:
        tmp = uzio[[UZIO_KEY, uzio_emp_status_col]].copy()
        tmp[uzio_emp_status_col] = tmp[uzio_emp_status_col].map(norm_blank)
        tmp = tmp[tmp[UZIO_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[UZIO_KEY]).strip()
            v = r[uzio_emp_status_col]
            if eid and norm_blank(v) != "" and eid not in uzio_status_map:
                uzio_status_map[eid] = str(v)

    paycom_status_map = {}
    if paycom_emp_status_col is not None:
        tmp = paycom[[PAYCOM_KEY, paycom_emp_status_col]].copy()
        tmp[paycom_emp_status_col] = tmp[paycom_emp_status_col].map(norm_blank)
        tmp = tmp[tmp[PAYCOM_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[PAYCOM_KEY]).strip()
            v = r[paycom_emp_status_col]
            if eid and norm_blank(v) != "" and eid not in paycom_status_map:
                paycom_status_map[eid] = str(v)

    def get_emp_status(eid: str) -> str:
        eid = (eid or "").strip()
        if eid in uzio_status_map:
            return str(uzio_status_map[eid])
        if eid in paycom_status_map:
            return str(paycom_status_map[eid])
        return ""

    # pay type map (prefer UZIO)
    uzio_pay_type_col = find_col(uzio.columns, "Pay Type")
    paycom_pay_type_col = find_col(paycom.columns, "Pay Type")

    pay_type_map = {}
    if uzio_pay_type_col is not None:
        tmp = uzio[[UZIO_KEY, uzio_pay_type_col]].copy()
        tmp[uzio_pay_type_col] = tmp[uzio_pay_type_col].map(norm_blank)
        tmp = tmp[tmp[UZIO_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[UZIO_KEY]).strip()
            v = r[uzio_pay_type_col]
            if eid and norm_blank(v) != "" and eid not in pay_type_map:
                pay_type_map[eid] = canonical_pay_type(v)

    if paycom_pay_type_col is not None:
        tmp = paycom[[PAYCOM_KEY, paycom_pay_type_col]].copy()
        tmp[paycom_pay_type_col] = tmp[paycom_pay_type_col].map(norm_blank)
        tmp = tmp[tmp[PAYCOM_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[PAYCOM_KEY]).strip()
            v = r[paycom_pay_type_col]
            if eid and norm_blank(v) != "" and eid not in pay_type_map:
                pay_type_map[eid] = canonical_pay_type(v)

    # index maps (keep first occurrence per employee)
    uzio_idx = {}
    for i, eid in uzio[UZIO_KEY].items():
        e = str(eid).strip()
        if e and e not in uzio_idx:
            uzio_idx[e] = i

    paycom_idx = {}
    for i, eid in paycom[PAYCOM_KEY].items():
        e = str(eid).strip()
        if e and e not in paycom_idx:
            paycom_idx[e] = i

    all_emps = sorted(set(uzio_idx.keys()).union(set(paycom_idx.keys())))

    rows = []
    for eid in all_emps:
        u_i = uzio_idx.get(eid)
        p_i = paycom_idx.get(eid)

        emp_status_context = get_emp_status(eid)
        emp_pay_type = pay_type_map.get(eid, "")

        for _, mr in mapping.iterrows():
            uz_field = mr["UZIO_Column"]
            pc_col = mr["PAYCOM_Resolved_Column"]

            uz_missing_row = (u_i is None)
            pc_missing_row = (p_i is None)

            uz_missing_col = (uz_field not in uzio.columns)
            pc_missing_col = (pc_col not in paycom.columns)

            uz_val = ""
            pc_val = ""
            if (not uz_missing_row) and (not uz_missing_col):
                uz_val = uzio.loc[u_i, uz_field]
            if (not pc_missing_row) and (not pc_missing_col):
                pc_val = paycom.loc[p_i, pc_col]

            # Decide status
            if pc_missing_row and (not uz_missing_row):
                status = "Employee ID Not Found in Paycom"
            elif uz_missing_row and (not pc_missing_row):
                status = "Employee ID Not Found in Uzio"
            elif pc_missing_col:
                status = "Column Missing in Paycom Sheet"
            elif uz_missing_col:
                status = "Column Missing in Uzio Sheet"
            else:
                # ✅ Pay-type based ignore rules (your latest requirement)
                if should_ignore_field_for_paytype(uz_field, emp_pay_type):
                    status = "Data Match"
                else:
                    same = normalized_compare(uz_field, uz_val, pc_val)
                    if same:
                        status = "Data Match"
                    else:
                        uz_b = norm_blank(uz_val)
                        pc_b = norm_blank(pc_val)
                        if (uz_b == "" or uz_b is None) and (pc_b != "" and pc_b is not None):
                            status = "Value missing in Uzio (Paycom has value)"
                        elif (uz_b != "" and uz_b is not None) and (pc_b == "" or pc_b is None):
                            status = "Value missing in Paycom (Uzio has value)"
                        else:
                            status = "Data Mismatch"

            rows.append(
                {
                    "Employee": eid,
                    "Field": uz_field,
                    "Employment Status": emp_status_context,  # extra context column
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
        "Data Match",
        "Data Mismatch",
        "Value missing in Uzio (Paycom has value)",
        "Value missing in Paycom (Uzio has value)",
        "Employee ID Not Found in Uzio",
        "Employee ID Not Found in Paycom",
        "Column Missing in Paycom Sheet",
        "Column Missing in Uzio Sheet",
    ]

    if not comparison_detail.empty:
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
    uzio_emps = set(uzio[UZIO_KEY].dropna().map(str))
    paycom_emps = set(paycom[PAYCOM_KEY].dropna().map(str))

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
                len(uzio_emps),
                len(paycom_emps),
                len(uzio_emps & paycom_emps),
                len(uzio_emps - paycom_emps),
                len(paycom_emps - uzio_emps),
                int(len(uzio)),
                int(len(paycom)),
                int(mapping.shape[0]),
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
        # requested format: Client_Name_Paycom_Census_Data_Audit_<Current Date>
        today_str = datetime.now().strftime("%Y-%m-%d")
        out_name = f"Client_Name_Paycom_Census_Data_Audit_{today_str}.xlsx"

        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
