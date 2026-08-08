# GL & TB Reconciler

The **GL & TB Reconciler** is a CPA-level internal auditing tool designed to automate the reconciliation process between General Ledger (GL) records and Trial Balance (TB) summaries. It also features a robust auto-mapping engine to instantly generate Financial Statements based on the reconciled data.

This repository contains two standalone versions:
1. **HTML/JS Version (Recommended for Web Browsers)**: Zero backend, runs 100% client-side directly in the browser via `index.html`.
2. **Python Version**: Streamlit-based Python web application located in the `python/` directory.

---

## Directory Structure

```
gl_tb_reconciler/
├── index.html            # Primary Web App (HTML version - double-click to open in browser)
├── app.js                # Full JavaScript application logic & Excel generation engine
├── styles.css            # Responsive dark/light theme styles
├── fs_config.json        # Financial statement mapping & structure configuration
├── FS.xlsx               # Sample / benchmark Excel template file
├── README.md             # Documentation
├── start.sh              # Root launcher script for the Python version
│
└── python/               # Dedicated Python (Streamlit) version folder
    ├── app.py            # Streamlit application source code
    ├── requirements.txt  # Python dependencies (Streamlit, Pandas, PyMuPDF, OpenPyXL)
    ├── fs_config.json    # Local configuration copy for Streamlit app
    └── start.sh          # Streamlit launcher script
```

---

## 1. Running the HTML/JS Version (No Installation Required)

Simply open `index.html` in any modern web browser:
- Double-click `index.html` or drag it into Chrome/Edge/Firefox/Safari.
- No Python, backend server, or Node.js required. Everything is processed client-side in the browser.

---

## 2. Running the Python Version

### Prerequisites
- Python 3.8+
- Virtual Environment

### Quick Start
To launch the Streamlit app from the root directory:
```bash
bash start.sh
```

Or navigate to the `python/` directory manually:
```bash
# 1. Create and activate a Python virtual environment
python3 -m venv venv
source venv/bin/activate

# 2. Install dependencies
pip install -r python/requirements.txt

# 3. Launch Streamlit
cd python
streamlit run app.py
```

---

## Key Features

- **Automated Discrepancy Detection**: Upload your TB (Excel) and GL (PDF) to instantly see mismatches, missing accounts, and matched items.
- **Brought Forward (ยอดยกมา) Verification**: Supports uploading a TB PDF to extract and cross-reference Brought Forward balances against GL.
- **Top 10 Suspects**: Automatically identifies high-balance accounts for priority manual review.
- **Intelligent Financial Statement Mapping**: Rule-based hybrid mapping (Account ID prefix and keyword matching) to map TB accounts to Financial Statements.
- **Dynamic Excel Export**: Generates full Reconciliation Reports and TFRS-compliant Financial Statements (`.xlsx`) directly in browser memory.
