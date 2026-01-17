# Standard Operating Procedure (SOP)
## Paycom vs. UZIO Census Audit Tool (Self-Service)

**Objective**: To automate the comparison of employee census data between UZIO and Paycom, identifying discrepancies in key fields for reconciliation.

---

### 1. Prerequisite: Prepare Input Data
You need a **single Excel Workbook (.xlsx)** containing exactly **three tabs** with the following specific names and structures:

#### Tab 1: `Uzio Data`
*   **Source**: Export relevant census data from UZIO.
*   **Required Column**: `Employee ID` (or `Employee Code`).
*   *Note*: Ensure headers are in the first row.

#### Tab 2: `Paycom Data`
*   **Source**: Export relevant census data from Paycom (Source of Truth).
*   **Required Column**: `Employee Code` (or `Employee ID`).
*   *Note*: This dataset is treated as the "Master" or "Source of Truth".

#### Tab 3: `Mapping Sheet`
*   **Purpose**: Tells the tool which UZIO column compares to which Paycom column.
*   **Required Columns**:
    *   `UZIO Column`: Exact header name from the *Uzio Data* tab.
    *   `Paycom Column`: Exact header name from the *Paycom Data* tab.
*   *Tip*: Only map the fields you want to audit.

---

### 2. Execution Steps
1.  **Launch the Tool**: Open the tool URL: [https://paycom-uzio-census-comparitor-tool.streamlit.app/](https://paycom-uzio-census-comparitor-tool.streamlit.app/)
2.  **Upload File**: Click **"Browse files"** and select your prepared Excel workbook.
    *   *System Check*: The 'Run Audit' button will become active once a valid file is uploaded.
3.  **Run Audit**: Click the **"Run Audit"** button.
    *   *Processing*: The system will normalize data (e.g., stripping 11-digit phone prefixes, matching "Full-Time" to "Full Time").
4.  **Download Results**: Once complete, a **"Download Report (.xlsx)"** button will appear. Click to save the audit file.

---

### 3. Understanding the Output Report
The downloaded Excel report contains three tabs:

*   **`Summary`**: High-level stats (Total Employees in each system, Overlap count).
*   **`Field_Summary_By_Status`**: Quick view of field health.
    *   **OK**: Data matches (using fuzzy logic).
    *   **MISMATCH**: Data differs between systems.
    *   **MISSING_IN_...**: Field is empty in one system but present in the other.
*   **`Comparison_Detail_AllFields`**: The master audit list.
    *   **Filter this tab** by the `PAYCOM_SourceOfTruth_Status` column to work through "MISMATCH" items.
    *   **Note**: "Employment Status" is provided for context (e.g., to see if mismatched employees are Terminated).
