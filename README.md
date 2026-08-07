# GL & TB Reconciler

The **GL & TB Reconciler** is a CPA-level internal auditing tool designed to automate the reconciliation process between General Ledger (GL) records and Trial Balance (TB) summaries. It also features a robust auto-mapping engine to instantly generate Financial Statements based on the reconciled data.

## Features

- **Automated Discrepancy Detection**: Upload your TB (Excel) and GL (PDF) to instantly see mismatches, missing accounts, and perfectly matched items.
- **Brought Forward (ยอดยกมา) Verification**: Supports uploading a TB PDF to extract and cross-reference the exact initial Brought Forward balances against the GL to ensure absolute accuracy.
- **Top 10 Suspects**: Automatically identifies and highlights the top 10 accounts with the highest absolute balances for priority manual review.
- **Intelligent Financial Statement Mapping**: Automatically maps Trial Balance accounts to their respective Financial Statement (FS) line items using a dual-priority approach (Account ID prefix and Keyword text matching).
- **Interactive Drill-down UI**: Review the FS mappings in an interactive table and click to expand the line items to see exactly which individual accounts make up the total.
- **Template Export**: Securely injects your finalized, calculated balances directly into an existing `FS.xlsx` template without modifying the original template file. Also exports the full reconciliation data (mismatches, matches, etc.) into a categorized Excel report.

## Prerequisites

- Python 3.8+
- Virtual Environment (recommended)

## Installation

1. Clone or download the repository to your local machine.
2. Create and activate a Python virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Ensure your template file named `FS.xlsx` is located in the parent directory (one folder up from `app.py`), or modify the path in `app.py` to point to your specific template location.

## Usage

1. Start the Streamlit application:
   ```bashหะ
   streamlit run app.py
   ```
2. Open your web browser to the local URL provided in the terminal (usually `http://localhost:8501`).
3. **Upload Files**:
   - Upload your Trial Balance in Excel format (`.xls` or `.xlsx`).
   - Upload your General Ledger in PDF format.
   - (Optional) Upload your Trial Balance in PDF format to verify Brought Forward balances.
4. Click **Run Reconciliation**.
5. **Review Results**: Use the tabs (`Discrepancies`, `Missing Records`, `Top 10 Suspects`, `All Matches`) to audit your data.
6. **FS Mapping**: Navigate to the `FS Mapping` tab to review the auto-guessed line items. Use the dropdowns in the data editor to manually correct any misclassifications.
7. **Export**: Go to the `Export Reports` tab to download your finalized Reconciliation Audit report or the Generated Financial Statements.

## Technologies Used

- **Streamlit**: Web interface and UI components.
- **Pandas**: Data extraction, merging, and numeric analysis.
- **PyMuPDF (fitz)**: Parsing and extracting structural text from complex accounting PDFs.
- **OpenPyXL & XlsxWriter**: Reading, writing, and formatting Excel templates.
