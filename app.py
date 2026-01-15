# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# Data_Audit_Tool — UZIO vs PAYCOM (PAYCOM = Source of Truth)
# Single Excel upload with 3 tabs:
#   1) "Uzio Data"
#   2) "Paycom Data"
#   3) "Mapping Sheet"
#
# Mapping Sheet columns (accepted):
#   - "Uzio Coloumn" / "Uzio Column"
#   - "Paycom Coloumn" / "Paycom Column"
#
# OUTPUT TABS (ONLY):
#   - Summary
#   - Field_Summary_By_Status   (columns G,H,I removed from this tab)
#   - Comparison_Detail_AllFields
#
# Status logic (PAYCOM is truth):
#   OK, MISMATCH, UZIO_MISSING_VALUE, PAYCOM_MISSING_VALUE,
#   MISSING_IN_UZIO, MISSING_IN_PAYCOM, PAYCOM_COLUMN_MISSING
#
# Fixes included (do not affect other logic):
# - Phone: digits-only normalization (ADP-style)
# - SSN/Tax ID: digits-only + pad to 9 (handles leading zeros)
# - Middle Initial vs Middle Name: "N" matches "Nicole"
# - Pay Type: "Salaried" matches "Salary" (hourly/salaried buckets)
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

UZIO_SHEET = "Uzio Data"
PAYCOM_SHEET = "Paycom Data"
MAP_SHEET = "Mapping Sheet"

# ---------- UI (no sidebar / no header/footer) ----------
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
st.write("Upload the Excel workbook (.xlsx) containing: Uzio Data, Paycom Data, and Mapping Sheet.")

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

def norm_phone_digits(x):
    """Digits-only phone normalize, drop leading 1, keep last 10 if longer."""
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

def normalize_middle_logic(field_name: str, uz_val, pay_val):
    """
    Middle initial vs middle name:
    - If UZIO looks like single letter and PAYCOM looks like a name,
      compare first letter only.
    """
    f = norm_colname(field_name).casefold()
    if "middle" in f and ("initial" in f or "middle name" in f):
        uz = norm_blank(uz_val)
        pv = norm_blank(pay_val)
        if uz == "" or pv == "":
            return None  # no override, normal pipeline handles missing
        uzs = re.sub(r"\s+", " ", str(uz).strip())
        pvs = re.sub(r"\s+", " ", str(pv).strip())
        # If UZIO is just one alphabet
        if re.fullmatch(r"[A-Za-z]", uzs) and len(pvs) >= 1:
            return uzs.casefold(), pvs[:1].casefold()
    return None

def normalize_paytype_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def paytype_bucket(paytype_norm: str) -> str:
    s = ("" if paytype_norm is None else str(paytype_norm)).casefold()
    if "hour" in s:
        return "hourly"
    if "salary" in s or "salaried" in s:
        return "salaried"
    return ""

NUMERIC_KEYWORDS = {"salary", "rate", "hours", "amount", "percent", "percentage"}
DATE_KEYWORDS = {"date", "dob", "birth", "effective"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
PHONE_KEYWORDS = {"phone", "mobile"}
SSN_KEYWORDS = {"ssn", "tax id", "taxid", "social security"}
PAYTYPE_KEYWORDS = {"pay type"}

def norm_value(x, field_name: str):
    f = norm_colname(field_name).casefold()
    x = norm_blank(x)
    if x == "":
        return ""

    if any(k in f for k in PHONE_KEYWORDS):
        return norm_phone_digits(x)

    if any(k in f for k in SSN_KEYWORDS):
        # pad to 9 so leading zeros match
        return digits_only_padded(x, 9)

    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    if any(k in f for k in PAYTYPE_KEYWORDS):
        # normalize buckets so "Salaried" == "Salary"
        b = paytype_bucket(normalize_paytype_text(x))
        return b if b else re.sub(r"\s+", " ", str(x).strip()).casefold()

    if any(k in f for k in NUMERIC_KEYWORDS):
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        if isinstance(x, str):
            s = x.strip().replace(",", "").replace("$", "")
            try:
                return float(s)
            except Exception:
                return re.sub(r"\s+", " ", x.strip()).casefold()

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

def get_mapping_cols(mapping_cols):
    """Accept both typo and correct spellings."""
    m = {norm_colname(c).casefold(): c for c in mapping_cols}
    uz = m.get("uzio coloumn") or m.get("uzio column")
    pc = m.get("paycom coloumn") or m.get("paycom column")
    return uz, pc

# ---------- Core comparison ----------
def run_comparison(file_bytes: bytes) -> bytes:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")

    # Read sheets
    uzio = pd.read_excel(xls, sheet_name=UZIO_SHEET, dtype=object)
    paycom = pd.read_excel(xls, sheet_name=PAYCOM_SHEET, dtype=object)
    mapping = pd.read_excel(xls, sheet_name=MAP_SHEET, dtype=object)

    # Normalize headers
    uzio.columns = [norm_colname(c) for c in uzio.columns]
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    mapping.columns = [norm_colname(c) for c in mapping.columns]

    uz_col_name, pc_col_name = get_mapping_cols(mapping.columns)
    if uz_col_name is None or pc_col_name is None:
        raise ValueError("Mapping Sheet must contain: 'Uzio Coloumn/Uzio Column' and 'Paycom Coloumn/Paycom Column'.")

    mapping[uz_col_name] = mapping[uz_col_name].map(norm_colname)
    mapping[pc_col_name] = mapping[pc_col_name].map(norm_colname)

    mapping_valid = mapping.dropna(subset=[uz_col_name, pc_col_name]).copy()
    mapping_valid = mapping_valid[(mapping_valid[uz_col_name] != "") & (mapping_valid[pc_col_name] != "")]
    mapping_valid = mapping_valid.drop_duplicates(subset=[uz_col_name], keep="first").copy()

    # Detect join key row
    key_row = mapping_valid[mapping_valid[uz_col_name].str.contains("Employee ID", case=False, na=False)]
    if len(key_row) == 0:
        raise ValueError("Mapping Sheet must include a row where UZIO column contains 'Employee ID' mapped to Paycom key column.")
    UZIO_KEY = key_row.iloc[0][uz_col_name]
    PAYCOM_KEY = key_row.iloc[0][pc_col_name]

    if UZIO_KEY not in uzio.columns:
        raise ValueError(f"UZIO key column '{UZIO_KEY}' not found in Uzio Data.")
    if PAYCOM_KEY not in paycom.columns:
        raise ValueError(f"Paycom key column '{PAYCOM_KEY}' not found in Paycom Data.")

    # Normalize keys
    uzio[UZIO_KEY] = norm_key_series(uzio[UZIO_KEY])
    paycom[PAYCOM_KEY] = norm_key_series(paycom[PAYCOM_KEY])

    uzio_keys = set(uzio[UZIO_KEY].dropna().astype(str).str.strip()) - {""}
    paycom_keys = set(paycom[PAYCOM_KEY].dropna().astype(str).str.strip()) - {""}
    all_keys = sorted(uzio_keys.union(paycom_keys))

    uzio_idx = uzio.set_index(UZIO_KEY, drop=False)
    paycom_idx = paycom.set_index(PAYCOM_KEY, drop=False)

    # Compare all mapped fields excluding key row
    mapped_fields = [f for f in mapping_valid[uz_col_name].tolist() if f != UZIO_KEY]
    uz_to_pc = dict(zip(mapping_valid[uz_col_name], mapping_valid[pc_col_name]))

    # Identify mappings where Paycom column missing (used only for status)
    paycom_missing_cols = set(mapping_valid.loc[~mapping_valid[pc_col_name].isin(paycom.columns), pc_col_name].tolist())

    rows = []
    for emp_id in all_keys:
        uz_exists = emp_id in uzio_idx.index
        pc_exists = emp_id in paycom_idx.index

        for field in mapped_fields:
            pc_col = uz_to_pc.get(field, "")

            uz_val = uzio_idx.at[emp_id, field] if uz_exists and field in uzio_idx.columns else ""
            pc_val = paycom_idx.at[emp_id, pc_col] if pc_exists and (pc_col in paycom_idx.columns) else ""

            # Status logic (PAYCOM truth)
            if not pc_exists and uz_exists:
                status = "MISSING_IN_PAYCOM"
            elif pc_exists and not uz_exists:
                status = "MISSING_IN_UZIO"
            elif pc_exists and uz_exists and (pc_col not in paycom.columns):
                status = "PAYCOM_COLUMN_MISSING"
            else:
                # Middle initial vs middle name special compare
                mid_override = normalize_middle_logic(field, uz_val, pc_val)
                if mid_override is not None:
                    uz_n, pc_n = mid_override
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
    comparison_detail = comparison_detail[["Employee ID", "Field", "UZIO_Value", "PAYCOM_Value", "PAYCOM_SourceOfTruth_Status"]]

    # ---------- Field_Summary_By_Status ----------
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

    pivot = pivot[statuses].copy()
    pivot["Total"] = pivot.sum(axis=1)
    pivot["OK"] = pivot["OK"].astype(int)
    pivot["NOT_OK"] = (pivot["Total"] - pivot["OK"]).astype(int)

    field_summary_by_status = pivot.reset_index()[
        ["Field", "Total", "OK", "NOT_OK"]
        + statuses[1:]  # everything except OK
    ]

    # Remove Excel columns G,H,I from Field_Summary_By_Status (by position)
    # (Excel columns are 1-indexed: A=1 ... G=7 H=8 I=9)
    # We drop dataframe columns at positions 6,7,8 (0-indexed) if present.
    cols = list(field_summary_by_status.columns)
    drop_idx = [6, 7, 8]
    drop_cols = [cols[i] for i in drop_idx if i < len(cols)]
    if drop_cols:
        field_summary_by_status = field_summary_by_status.drop(columns=drop_cols)

    # ---------- Summary ----------
    summary = pd.DataFrame(
        {
            "Metric": [
                "Employees in UZIO sheet",
                "Employees in PAYCOM sheet",
                "Employees present in both",
                "Employees missing in PAYCOM (UZIO only)",
                "Employees missing in UZIO (PAYCOM only)",
                "Mapped fields total (from mapping sheet)",
                "Total comparison rows (employees x mapped fields)",
            ],
            "Value": [
                len(uzio_keys),
                len(paycom_keys),
                len(uzio_keys.intersection(paycom_keys)),
                len(uzio_keys - paycom_keys),
                len(paycom_keys - uzio_keys),
                len(mapped_fields),
                comparison_detail.shape[0],
            ],
        }
    )

    # ---------- Export report ----------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Field_Summary_By_Status", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison_Detail_AllFields", index=False)

    return out.getvalue()

# ---------- UI ----------
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
