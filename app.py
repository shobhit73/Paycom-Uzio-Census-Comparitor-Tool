import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom vs UZIO Census Audit Tool (PAYCOM = Source of Truth)
# INPUT workbook tabs (single upload):
#   - Uzio Data
#   - Paycom Data
#   - Mapping  (or "Mapping Sheet")
#
# OUTPUT workbook tabs:
#   - Summary
#   - Field_Summary_By_Status
#   - Comparison_Detail_AllFields
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET_CANDIDATES = ["Uzio Data", "UZIO Data", "Uzio", "UZIO"]
PAYCOM_SHEET_CANDIDATES = ["Paycom Data", "PAYCOM Data", "Paycom", "PAYCOM"]
MAP_SHEET_CANDIDATES = ["Mapping", "Mapping Sheet", "MAP", "Map"]

# ---------- UI config (no sidebar / no previews) ----------
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

st.title(APP_TITLE)

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

def digits_only(x):
    x = norm_blank(x)
    if x == "":
        return ""
    try:
        if isinstance(x, (int, np.integer)):
            return str(int(x))
        if isinstance(x, (float, np.floating)) and float(x).is_integer():
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
    """
    Make these equivalent:
      Jr. == JR == jr
    """
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).strip()
    s = re.sub(r"[^A-Za-z0-9]", "", s)  # remove '.' etc.
    return s.casefold()

def normalize_reason_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = s.replace("’", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip('"').strip("'")
    return s.casefold()

def normalize_paytype_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    return s

def paytype_bucket(paytype_norm: str) -> str:
    s = ("" if paytype_norm is None else str(paytype_norm)).casefold()
    if "hour" in s:
        return "hourly"
    if "salary" in s or "salaried" in s:
        return "salaried"
    return ""

def is_annual_salary_field(field_name: str) -> bool:
    return "annual salary" in norm_colname(field_name).casefold()

def is_hourly_rate_field(field_name: str) -> bool:
    f = norm_colname(field_name).casefold()
    return ("hourly pay rate" in f) or ("hourly rate" in f)

def is_termination_reason_field(field_name: str) -> bool:
    return "termination reason" in norm_colname(field_name).casefold()

def contains_word(s: str, word: str) -> bool:
    s = ("" if s is None else str(s)).casefold()
    return word.casefold() in s

# Paycom -> UZIO "Other" acceptable reasons (from your list)
ALLOWED_PAYCOM_REASONS_AS_OTHER = {
    "attendance violation",
    "employee elected not to return",
    "failure to show up for work",
    "poor performance and attendance",
    "attendance violation and poor performance",
    "never worked",
    "performance, safety, and attendance violations",
    "left for military service",
    "policy violation",
    "personal issue",
    "poor performance, poor attitude and attendance violations",
    "health issue",
    "conflict with other job",
    "poor performance",
    "did not start work",
    "poor attitude and performance",
    "poor attendance, attitude, and performance.",
    "professionalism, performance and attendance issues",
}

NUMERIC_KEYWORDS = {"salary", "rate", "hours", "amount"}
DATE_KEYWORDS = {"date", "dob", "birth", "doh"}  # <-- so Original DOH is treated as date
SSN_KEYWORDS = {"ssn", "tax id"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
PHONE_KEYWORDS = {"phone", "mobile"}
SUFFIX_KEYWORDS = {"suffix"}
PAYTYPE_KEYWORDS = {"pay type"}

def norm_value(x, field_name: str):
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    # Suffix normalization (Jr. == JR)
    if any(k in f for k in SUFFIX_KEYWORDS):
        return norm_suffix(x)

    # Phone normalization (digits only)
    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)

    # SSN normalization (digits only, pad to 9 so leading zeros are preserved)
    if any(k in f for k in SSN_KEYWORDS):
        return digits_only_padded(x, 9)

    # ZIP normalization
    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    # Date normalization (Original DOH etc.)
    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    # Pay Type normalization (Salaried == Salary)
    if any(k in f for k in PAYTYPE_KEYWORDS):
        b = paytype_bucket(normalize_paytype_text(x))
        return b  # "hourly" / "salaried" / ""

    # Numeric normalization
    if any(k in f for k in NUMERIC_KEYWORDS):
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        if isinstance(x, str):
            s = x.strip().replace(",", "").replace("$", "")
            try:
                return float(s)
            except Exception:
                return re.sub(r"\s+", " ", x.strip()).casefold()

    # Default string normalization
    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()

    return str(x).casefold()

def norm_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

def pick_sheet(xls: pd.ExcelFile, candidates: list[str]) -> str:
    existing = {name.casefold(): name for name in xls.sheet_names}
    for c in candidates:
        if c.casefold() in existing:
            return existing[c.casefold()]
    raise ValueError(f"Missing required sheet. Expected one of: {candidates}")

def resolve_col_label(label: str, df_cols: list[str]) -> str:
    """
    Resolve mapping labels to actual dataframe column names (case/space/punctuation tolerant).
    """
    if label is None:
        return ""
    raw = norm_colname(label)
    if raw == "":
        return ""

    col_norm = {norm_colname(c).casefold(): c for c in df_cols}

    k = raw.casefold()
    if k in col_norm:
        return col_norm[k]

    # try loose contains match
    for kn, actual in col_norm.items():
        if kn and (kn in k or k in kn):
            return actual

    return ""

# ---------- Core comparison ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uz_sheet = pick_sheet(xls, UZIO_SHEET_CANDIDATES)
    pc_sheet = pick_sheet(xls, PAYCOM_SHEET_CANDIDATES)
    map_sheet = pick_sheet(xls, MAP_SHEET_CANDIDATES)

    uzio = pd.read_excel(xls, sheet_name=uz_sheet, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=pc_sheet, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=map_sheet, dtype=object)

    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    # Mapping column headers (support "Coloumn" typo)
    uz_map_col = None
    pc_map_col = None
    for c in mapping.columns:
        cc = norm_colname(c).casefold()
        if cc in {"uzio coloumn", "uzio column"}:
            uz_map_col = c
        if cc in {"paycom coloumn", "paycom column"}:
            pc_map_col = c

    if uz_map_col is None or pc_map_col is None:
        raise ValueError("Mapping sheet must contain columns: 'Uzio Coloumn/Column' and 'Paycom Coloumn/Column'.")

    mapping[uz_map_col] = mapping[uz_map_col].map(norm_colname)
    mapping[pc_map_col] = mapping[pc_map_col].map(norm_colname)

    mapping_valid = mapping.dropna(subset=[uz_map_col, pc_map_col]).copy()
    mapping_valid = mapping_valid[(mapping_valid[uz_map_col] != "") & (mapping_valid[pc_map_col] != "")]
    mapping_valid = mapping_valid.drop_duplicates(subset=[uz_map_col], keep="first").copy()

    # Identify key row (Employee ID mapping)
    key_row = mapping_valid[mapping_valid[uz_map_col].str.contains("Employee ID", case=False, na=False)]
    if len(key_row) == 0:
        raise ValueError("Mapping must include UZIO 'Employee ID' mapped to PAYCOM employee key column.")
    UZIO_KEY_LABEL = key_row.iloc[0][uz_map_col]
    PAYCOM_KEY_LABEL = key_row.iloc[0][pc_map_col]

    UZIO_KEY = resolve_col_label(UZIO_KEY_LABEL, list(uzio.columns))
    PAYCOM_KEY = resolve_col_label(PAYCOM_KEY_LABEL, list(paycom.columns))

    if UZIO_KEY == "" or UZIO_KEY not in uzio.columns:
        raise ValueError(f"UZIO key column '{UZIO_KEY_LABEL}' not found in Uzio Data.")
    if PAYCOM_KEY == "" or PAYCOM_KEY not in paycom.columns:
        raise ValueError(f"PAYCOM key column '{PAYCOM_KEY_LABEL}' not found in Paycom Data.")

    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    uzio_idx = uzio.set_index(UZIO_KEY, drop=False)
    paycom_idx = paycom.set_index(PAYCOM_KEY, drop=False)

    uz_keys = set(uzio[UZIO_KEY].dropna().astype(str).str.strip()) - {""}
    pc_keys = set(paycom[PAYCOM_KEY].dropna().astype(str).str.strip()) - {""}
    all_keys = sorted(uz_keys.union(pc_keys))

    # Build field mapping dict (UZIO field -> PAYCOM column resolved)
    uz_fields_all = mapping_valid[uz_map_col].tolist()
    mapped_fields = [f for f in uz_fields_all if norm_colname(f).casefold() != norm_colname(UZIO_KEY_LABEL).casefold()]

    uz_to_pc_label = dict(zip(mapping_valid[uz_map_col], mapping_valid[pc_map_col]))
    uz_to_pc_resolved = {}
    for uz_field, pc_label in uz_to_pc_label.items():
        uz_to_pc_resolved[uz_field] = resolve_col_label(pc_label, list(paycom.columns))

    # Pay Type columns (for exceptions)
    paytype_row = mapping_valid[mapping_valid[uz_map_col].str.contains(r"\bpay\s*type\b", case=False, na=False)]
    UZIO_PAYTYPE_FIELD = paytype_row.iloc[0][uz_map_col] if len(paytype_row) else None
    PAYCOM_PAYTYPE_COL = resolve_col_label(paytype_row.iloc[0][pc_map_col], list(paycom.columns)) if len(paytype_row) else None

    def get_employee_paytype(emp_id: str, uz_exists: bool, pc_exists: bool) -> str:
        # Prefer PAYCOM (source of truth)
        if pc_exists and PAYCOM_PAYTYPE_COL and PAYCOM_PAYTYPE_COL in paycom_idx.columns:
            v = paycom_idx.at[emp_id, PAYCOM_PAYTYPE_COL]
            if norm_blank(v) != "":
                return str(v)
        if uz_exists and UZIO_PAYTYPE_FIELD and UZIO_PAYTYPE_FIELD in uzio_idx.columns:
            v = uzio_idx.at[emp_id, UZIO_PAYTYPE_FIELD]
            if norm_blank(v) != "":
                return str(v)
        return ""

    rows = []
    for emp_id in all_keys:
        uz_exists = emp_id in uzio_idx.index
        pc_exists = emp_id in paycom_idx.index

        emp_paytype = get_employee_paytype(emp_id, uz_exists=uz_exists, pc_exists=pc_exists)
        emp_pay_bucket = paytype_bucket(normalize_paytype_text(emp_paytype))

        for field in mapped_fields:
            pc_col = uz_to_pc_resolved.get(field, "")

            uz_val = uzio_idx.at[emp_id, field] if (uz_exists and field in uzio_idx.columns) else ""
            pc_val = paycom_idx.at[emp_id, pc_col] if (pc_exists and pc_col in paycom_idx.columns) else ""

            # Statusing (PAYCOM = source of truth)
            if not pc_exists and uz_exists:
                status = "MISSING_IN_PAYCOM"
            elif pc_exists and not uz_exists:
                status = "MISSING_IN_UZIO"
            elif pc_exists and uz_exists and (pc_col == "" or pc_col not in paycom.columns):
                status = "PAYCOM_COLUMN_MISSING"
            elif pc_exists and uz_exists and (field not in uzio.columns):
                status = "UZIO_COLUMN_MISSING"
            else:
                uz_n = norm_value(uz_val, field)
                pc_n = norm_value(pc_val, field)

                # --- Termination Reason special handling ---
                if is_termination_reason_field(field):
                    uz_reason = normalize_reason_text(uz_val)
                    pc_reason = normalize_reason_text(pc_val)

                    # Rule A: If both contain "voluntary" => OK; both contain "involuntary" => OK
                    if (contains_word(uz_reason, "voluntary") and contains_word(pc_reason, "voluntary")):
                        status = "OK"
                    elif (contains_word(uz_reason, "involuntary") and contains_word(pc_reason, "involuntary")):
                        status = "OK"
                    # Rule B: Paycom reasons that map to UZIO "Other" => OK
                    elif uz_reason == "other" and pc_reason in ALLOWED_PAYCOM_REASONS_AS_OTHER:
                        status = "OK"
                    else:
                        if (uz_n == pc_n) or (uz_n == "" and pc_n == ""):
                            status = "OK"
                        elif uz_n == "" and pc_n != "":
                            status = "UZIO_MISSING_VALUE"
                        elif uz_n != "" and pc_n == "":
                            status = "PAYCOM_MISSING_VALUE"
                        else:
                            status = "MISMATCH"
                else:
                    if (uz_n == pc_n) or (uz_n == "" and pc_n == ""):
                        status = "OK"
                    elif uz_n == "" and pc_n != "":
                        status = "UZIO_MISSING_VALUE"
                    elif uz_n != "" and pc_n == "":
                        status = "PAYCOM_MISSING_VALUE"
                    else:
                        status = "MISMATCH"

                # --- PayType exceptions (same behavior as your ADP logic) ---
                if status == "UZIO_MISSING_VALUE":
                    if emp_pay_bucket == "hourly" and is_annual_salary_field(field):
                        status = "OK"
                    elif emp_pay_bucket == "salaried" and is_hourly_rate_field(field):
                        status = "OK"

            rows.append(
                {
                    "Employee ID": emp_id,
                    "Field": field,
                    "UZIO_Value": uz_val,
                    "PAYCOM_Value": pc_val,
                    "PAYCOM_SourceOfTruth_Status": status,
                }
            )

    comparison_detail = pd.DataFrame(rows, columns=[
        "Employee ID", "Field", "UZIO_Value", "PAYCOM_Value", "PAYCOM_SourceOfTruth_Status"
    ])

    # Field Summary
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

    pivot = comparison_detail.pivot_table(
        index="Field",
        columns="PAYCOM_SourceOfTruth_Status",
        values="Employee ID",
        aggfunc="count",
        fill_value=0
    )

    for c in statuses:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["Total"] = pivot[statuses].sum(axis=1)
    pivot["OK"] = pivot["OK"].astype(int)
    pivot["NOT_OK"] = (pivot["Total"] - pivot["OK"]).astype(int)

    field_summary_by_status = pivot.reset_index()[[
        "Field",
        "Total",
        "OK",
        "NOT_OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "PAYCOM_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_PAYCOM",
        "PAYCOM_COLUMN_MISSING",
        "UZIO_COLUMN_MISSING",
    ]]

    summary = pd.DataFrame({
        "Metric": [
            "Employees in UZIO sheet",
            "Employees in PAYCOM sheet",
            "Employees present in both",
            "Employees missing in PAYCOM (UZIO only)",
            "Employees missing in UZIO (PAYCOM only)",
            "Mapped fields total (from mapping sheet)",
            "Total comparison rows (employees x mapped fields)",
            "Total NOT OK rows"
        ],
        "Value": [
            len(uz_keys),
            len(pc_keys),
            len(uz_keys.intersection(pc_keys)),
            len(uz_keys - pc_keys),
            len(pc_keys - uz_keys),
            len(mapped_fields),
            comparison_detail.shape[0],
            int((comparison_detail["PAYCOM_SourceOfTruth_Status"] != "OK").sum())
        ]
    })

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Field_Summary_By_Status", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison_Detail_AllFields", index=False)

    return out.getvalue()

# ---------- UI ----------
st.write("Upload the Excel workbook (.xlsx) containing: Uzio Data, Paycom Data, and Mapping (or Mapping Sheet).")
uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
run_btn = st.button("Run Audit", type="primary", disabled=(uploaded_file is None))

if run_btn:
    try:
        with st.spinner("Running audit..."):
            report_bytes = run_comparison(uploaded_file.getvalue())

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"UZIO_vs_PAYCOM_Comparison_Report_PAYCOM_SourceOfTruth_{ts}.xlsx"

        st.success("Report generated.")
        st.download_button(
            label="Download Report (.xlsx)",
            data=report_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
