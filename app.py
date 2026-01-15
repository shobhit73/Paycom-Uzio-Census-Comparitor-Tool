# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Paycom vs UZIO – Census Audit Tool (PAYCOM = Source of Truth)
# INPUT (single Excel workbook with 3 tabs):
#   - Uzio Data
#   - Paycom Data
#   - Mapping   (or "Mapping Sheet")
#
# OUTPUT tabs:
#   - Summary
#   - Field_Summary_By_Status
#   - Comparison_Detail_AllFields
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET_CANDIDATES = ["Uzio Data", "UZIO Data", "UZIO", "Uzio"]
PAYCOM_SHEET_CANDIDATES = ["Paycom Data", "PAYCOM Data", "PAYCOM", "Paycom"]
MAP_SHEET_CANDIDATES = ["Mapping", "Mapping Sheet", "MAP", "Map"]

# ---------- UI ----------
st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
st.markdown(
    """
    <style>
      [data-testid="stSidebar"] { display: none !important; }
      [data-testid="collapsedControl"] { display: none !important; }
      header { display: none !important; }
      footer { display: none !important; }
      .block-container { padding-top: 1.5rem; }
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

def norm_suffix(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).strip().replace("\u00A0", " ")
    s = re.sub(r"[^A-Za-z0-9]", "", s)  # remove punctuation (e.g., Jr. -> Jr)
    return s.casefold()

NUMERIC_KEYWORDS = {"salary", "rate", "hours", "amount", "percent", "percentage"}
DATE_KEYWORDS = {"date", "dob", "birth", "effective", "doh"}  # includes DOH so Original DOH matches
SSN_KEYWORDS = {"ssn", "tax id"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
PHONE_KEYWORDS = {"phone", "mobile"}
SUFFIX_KEYWORDS = {"suffix"}

def normalize_text_basic(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()
    return str(x).strip().casefold()

# Paycom Termination Reason -> allowed UZIO Termination Reason values (normalized)
TERM_REASON_ALLOWED = {
    # Paycom reason -> UZIO "Other"
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
    # Paycom reason -> multiple acceptable UZIO values
    "voluntary termination of employment": {
        "voluntary termination",
        "voluntary resignation",
        "voluntary",
    },
    "involuntary termination of employment": {
        "involuntary termination",
    },
}

def is_termination_reason_field(field_name: str) -> bool:
    return "termination reason" in norm_colname(field_name).casefold()

def norm_value(x, field_name: str):
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    if any(k in f for k in SUFFIX_KEYWORDS):
        return norm_suffix(x)

    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)

    if any(k in f for k in SSN_KEYWORDS):
        # pad to 9 so Excel numeric formatting doesn't drop leading zeros
        return digits_only_padded(x, 9)

    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    if any(k in f for k in NUMERIC_KEYWORDS):
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        if isinstance(x, str):
            s = x.strip().replace(",", "").replace("$", "")
            try:
                return float(s)
            except Exception:
                return normalize_text_basic(x)

    return normalize_text_basic(x)

def norm_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

def find_sheet_name(xls: pd.ExcelFile, candidates):
    available = set(xls.sheet_names)
    for c in candidates:
        if c in available:
            return c
    # fallback: case-insensitive match
    low_map = {s.casefold(): s for s in xls.sheet_names}
    for c in candidates:
        if c.casefold() in low_map:
            return low_map[c.casefold()]
    return None

def find_mapping_columns(mapping_df: pd.DataFrame):
    cols = list(mapping_df.columns)
    normed = {norm_colname(c).casefold(): c for c in cols}

    uz = None
    pc = None

    # Prefer exact common spellings
    for k in ["uzio coloumn", "uzio column"]:
        if k in normed:
            uz = normed[k]
            break

    for k in ["paycom coloumn", "paycom column"]:
        if k in normed:
            pc = normed[k]
            break

    # Fallback: fuzzy contains
    if uz is None:
        for c in cols:
            cc = norm_colname(c).casefold()
            if "uzio" in cc and ("coloumn" in cc or "column" in cc):
                uz = c
                break

    if pc is None:
        for c in cols:
            cc = norm_colname(c).casefold()
            if "paycom" in cc and ("coloumn" in cc or "column" in cc):
                pc = c
                break

    return uz, pc

# ---------- Core comparison ----------
def run_census_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    uzio_sheet = find_sheet_name(xls, UZIO_SHEET_CANDIDATES)
    paycom_sheet = find_sheet_name(xls, PAYCOM_SHEET_CANDIDATES)
    map_sheet = find_sheet_name(xls, MAP_SHEET_CANDIDATES)

    if uzio_sheet is None:
        raise ValueError(f"Could not find UZIO sheet. Expected one of: {UZIO_SHEET_CANDIDATES}")
    if paycom_sheet is None:
        raise ValueError(f"Could not find Paycom sheet. Expected one of: {PAYCOM_SHEET_CANDIDATES}")
    if map_sheet is None:
        raise ValueError(f"Could not find Mapping sheet. Expected one of: {MAP_SHEET_CANDIDATES}")

    uzio = pd.read_excel(xls, sheet_name=uzio_sheet, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=paycom_sheet, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=map_sheet, dtype=object)

    # Normalize headers
    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    uz_col, pc_col = find_mapping_columns(mapping)
    if uz_col is None or pc_col is None:
        raise ValueError("Mapping sheet must contain columns like 'Uzio Coloumn' and 'Paycom Coloumn' (or '... Column').")

    mapping[uz_col] = mapping[uz_col].map(norm_colname)
    mapping[pc_col] = mapping[pc_col].map(norm_colname)

    mapping_valid = mapping.dropna(subset=[uz_col, pc_col]).copy()
    mapping_valid = mapping_valid[(mapping_valid[uz_col] != "") & (mapping_valid[pc_col] != "")]
    mapping_valid = mapping_valid.drop_duplicates(subset=[uz_col], keep="first").copy()

    # Determine join key from mapping (UZIO Employee ID -> Paycom key)
    key_row = mapping_valid[mapping_valid[uz_col].str.contains("Employee ID", case=False, na=False)]
    if len(key_row) == 0:
        raise ValueError("Mapping must include a row where UZIO column contains 'Employee ID' mapped to Paycom key.")
    UZIO_KEY = key_row.iloc[0][uz_col]
    PAYCOM_KEY = key_row.iloc[0][pc_col]

    if UZIO_KEY not in uzio.columns:
        raise ValueError(f"UZIO key column '{UZIO_KEY}' not found in UZIO sheet.")
    if PAYCOM_KEY not in paycom.columns:
        raise ValueError(f"Paycom key column '{PAYCOM_KEY}' not found in Paycom sheet.")

    # Normalize keys
    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    uzio_keys = set(uzio[UZIO_KEY].dropna().astype(str).str.strip()) - {""}
    paycom_keys = set(paycom[PAYCOM_KEY].dropna().astype(str).str.strip()) - {""}
    all_keys = sorted(uzio_keys.union(paycom_keys))

    uzio_idx = uzio.set_index(UZIO_KEY, drop=False)
    paycom_idx = paycom.set_index(PAYCOM_KEY, drop=False)

    # Map UZIO field -> Paycom column
    uz_to_pc = dict(zip(mapping_valid[uz_col], mapping_valid[pc_col]))
    mapped_fields = [f for f in mapping_valid[uz_col].tolist() if f != UZIO_KEY]

    rows = []
    for emp_id in all_keys:
        uz_exists = emp_id in uzio_idx.index
        pc_exists = emp_id in paycom_idx.index

        for field in mapped_fields:
            pc_field = uz_to_pc.get(field, "")

            uz_val = uzio_idx.at[emp_id, field] if uz_exists and field in uzio_idx.columns else ""
            pc_val = paycom_idx.at[emp_id, pc_field] if pc_exists and (pc_field in paycom_idx.columns) else ""

            # Status logic (PAYCOM is source of truth)
            if not pc_exists and uz_exists:
                status = "MISSING_IN_PAYCOM"
            elif pc_exists and not uz_exists:
                status = "MISSING_IN_UZIO"
            elif pc_exists and uz_exists and (pc_field not in paycom.columns):
                status = "PAYCOM_COLUMN_MISSING"
            else:
                # Special rule: Termination Reason mapping
                if is_termination_reason_field(field):
                    uz_norm = normalize_text_basic(uz_val)
                    pc_norm = normalize_text_basic(pc_val)
                    allowed = TERM_REASON_ALLOWED.get(pc_norm)
                    if allowed is not None and uz_norm in allowed:
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

            rows.append(
                {
                    "Employee ID": emp_id,
                    "Field": field,
                    "UZIO_Value": uz_val,
                    "PAYCOM_Value": pc_val,
                    "PAYCOM_SourceOfTruth_Status": status,
                }
            )

    comparison_detail = pd.DataFrame(rows)
    comparison_detail = comparison_detail[
        ["Employee ID", "Field", "UZIO_Value", "PAYCOM_Value", "PAYCOM_SourceOfTruth_Status"]
    ]

    # Field summary by status
    statuses = [
        "OK",
        "MISMATCH",
        "UZIO_MISSING_VALUE",
        "PAYCOM_MISSING_VALUE",
        "MISSING_IN_UZIO",
        "MISSING_IN_PAYCOM",
        "PAYCOM_COLUMN_MISSING",
    ]
    pivot = comparison_detail.pivot_table(
        index="Field",
        columns="PAYCOM_SourceOfTruth_Status",
        values="Employee ID",
        aggfunc="count",
        fill_value=0,
    )
    for s in statuses:
        if s not in pivot.columns:
            pivot[s] = 0
    pivot = pivot[statuses]
    pivot["Total"] = pivot.sum(axis=1)
    pivot["OK"] = pivot["OK"].astype(int)
    pivot["NOT_OK"] = (pivot["Total"] - pivot["OK"]).astype(int)

    field_summary_by_status = pivot.reset_index()[
        ["Field", "Total", "OK", "NOT_OK", "MISMATCH", "UZIO_MISSING_VALUE", "PAYCOM_MISSING_VALUE",
         "MISSING_IN_UZIO", "MISSING_IN_PAYCOM", "PAYCOM_COLUMN_MISSING"]
    ]

    # Summary metrics
    summary = pd.DataFrame(
        {
            "Metric": [
                "Employees in UZIO sheet",
                "Employees in Paycom sheet",
                "Employees present in both",
                "Employees missing in Paycom (UZIO only)",
                "Employees missing in UZIO (Paycom only)",
                "Mapped fields total (from mapping sheet)",
                "Total comparison rows (employees x mapped fields)",
                "Total NOT OK rows",
            ],
            "Value": [
                len(uzio_keys),
                len(paycom_keys),
                len(uzio_keys.intersection(paycom_keys)),
                len(uzio_keys - paycom_keys),
                len(paycom_keys - uzio_keys),
                len(mapped_fields),
                int(comparison_detail.shape[0]),
                int((comparison_detail["PAYCOM_SourceOfTruth_Status"] != "OK").sum()),
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
st.write("Upload the Excel workbook (.xlsx) with 3 tabs: Uzio Data, Paycom Data, and Mapping. Then download the audit report.")

uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
run_btn = st.button("Run Audit", type="primary", disabled=(uploaded_file is None))

if run_btn:
    try:
        with st.spinner("Running audit..."):
            report_bytes = run_census_comparison(uploaded_file.getvalue())

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
