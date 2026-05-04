import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
import math
import openpyxl
import os

FS_LINE_ITEMS = {
    "เงินสดและรายการเทียบเท่าเงินสด": {"prefixes": ["111"], "keywords": ["เงินสด", "เงินฝาก"]},
    "ลูกหนี้การค้า": {"prefixes": ["113"], "keywords": ["ลูกหนี้"]},
    "สินทรัพย์หมุนเวียนอื่น ": {"prefixes": ["115", "119"], "keywords": ["ภาษีถูกหัก", "จ่ายล่วงหน้า"]},
    "อุปกรณ์-สุทธิ": {"prefixes": ["141", "142"], "keywords": ["เครื่องมือ", "เครื่องจักร", "อุปกรณ์", "ค่าเสื่อม"]},
    "เจ้าหนี้อื่น": {"prefixes": ["211"], "keywords": ["เจ้าหนี้", "กรมสรรพากร"]},
    "หนี้สินหมุนเวียนอื่น ": {"prefixes": ["213"], "keywords": ["ภาษีขาย", "ภาษีหัก", "ค้างจ่าย"]},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"prefixes": ["2138"], "keywords": ["เงินกู้ยืม"]},
    "ทุนเรือนหุ้น ": {"prefixes": ["31"], "keywords": ["ทุน"]},
    "กำไร(ขาดทุน)สะสม": {"prefixes": ["32"], "keywords": ["กำไร"]},
    "รายได้จากการให้บริการ": {"prefixes": ["41"], "keywords": ["รายได้จากการ"]},
    "รายได้อื่น": {"prefixes": ["42"], "keywords": ["รายได้อื่น", "ดอกเบี้ย"]},
    "ต้นทุนการให้บริการ": {"prefixes": ["51"], "keywords": ["ต้นทุน", "ซื้อ"]},
    "ค่าใช้จ่ายในการขายและบริหาร": {"prefixes": ["52", "53"], "keywords": ["ค่าใช้จ่าย", "เงินเดือน", "ค่าธรรมเนียม", "ค่าเสื่อม", "ค่าเช่า", "ประกัน", "สอบบัญชี", "บริการ"]},
}

FS_CELLS = {
    "เงินสดและรายการเทียบเท่าเงินสด": {"sheet": "สินทรัพย์", "cell": "C9"},
    "ลูกหนี้การค้า": {"sheet": "สินทรัพย์", "cell": "C10"},
    "สินทรัพย์หมุนเวียนอื่น ": {"sheet": "สินทรัพย์", "cell": "C11"},
    "อุปกรณ์-สุทธิ": {"sheet": "สินทรัพย์", "cell": "C15"},
    "เจ้าหนี้อื่น": {"sheet": "สินทรัพย์", "cell": "C42"},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"sheet": "สินทรัพย์", "cell": "C43"},
    "หนี้สินหมุนเวียนอื่น ": {"sheet": "สินทรัพย์", "cell": "C44"},
    "รายได้จากการให้บริการ": {"sheet": "กำไรขาดทุน", "cell": "D8"},
    "รายได้อื่น": {"sheet": "กำไรขาดทุน", "cell": "D9"},
    "ต้นทุนการให้บริการ": {"sheet": "กำไรขาดทุน", "cell": "D13"},
    "ค่าใช้จ่ายในการขายและบริหาร": {"sheet": "กำไรขาดทุน", "cell": "D14"},
}

def auto_map(row):
    acc_id = str(row.get('Account ID', ''))
    acc_name = str(row.get('Account Name', ''))
    
    matches = []
    for fs_line, rules in FS_LINE_ITEMS.items():
        for prefix in rules["prefixes"]:
            if acc_id.startswith(prefix):
                matches.append((len(prefix), fs_line))
    
    if matches:
        matches.sort(key=lambda x: x[0], reverse=True)
        return matches[0][1]
        
    for fs_line, rules in FS_LINE_ITEMS.items():
        for keyword in rules["keywords"]:
            if keyword in acc_name:
                return fs_line
                
    return "ไม่จัดประเภท (Unmapped)"


def parse_tb_pdf(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
        
    lines = text.split('\n')
    tb_bf_data = []
    
    for line in lines:
        line = line.strip()
        # Look for account ID at start
        match_id = re.match(r'^(\d{4}-\d{2})\s+(.*)', line)
        if match_id:
            acc_id = match_id.group(1)
            rest_of_line = match_id.group(2)
            
            numbers = re.findall(r'((?:\()?[\d,]+\.\d{2}(?:\))?)', rest_of_line)
            if len(numbers) >= 6:
                try:
                    bf_dr_str = numbers[-6].replace(',', '').replace('(', '-').replace(')', '')
                    bf_cr_str = numbers[-5].replace(',', '').replace('(', '-').replace(')', '')
                    
                    bf_dr = float(bf_dr_str)
                    bf_cr = float(bf_cr_str)
                    
                    tb_bf_data.append({
                        'Account ID': acc_id,
                        'TB Brought Forward Net': abs(bf_dr - bf_cr)
                    })
                except Exception:
                    pass
                    
    return pd.DataFrame(tb_bf_data)

st.set_page_config(page_title="GL & TB Reconciler", layout="wide")

def parse_tb(excel_file):
    # Read the excel file, skipping the first 4 rows to get to the data
    df = pd.read_excel(excel_file, header=None)
    
    # Looking for the data starting after header rows.
    # From previous check, row 3 has 'เลขที่บัญชี', row 4 has 'เดบิต' / 'เครดิต', row 5 is the first data row.
    
    # We will just iterate through rows and find where 'เลขที่บัญชี' style data starts
    tb_data = []
    start_reading = False
    
    for index, row in df.iterrows():
        acc_id = str(row[0]).strip()
        if not start_reading:
            if re.match(r'^\d{4}-\d{2}$', acc_id):
                start_reading = True
            else:
                continue
                
        if start_reading:
            if pd.isna(row[0]) or str(row[0]).strip() == 'nan':
                continue # Skip empty rows
            
            acc_id = str(row[0]).strip()
            if not re.match(r'^\d{4}-\d{2}$', acc_id):
                continue # Skip non-account rows
                
            acc_name = str(row[1]).strip()
            debit = float(row[2]) if not pd.isna(row[2]) else 0.0
            credit = float(row[3]) if not pd.isna(row[3]) else 0.0
            
            net_balance = abs(debit - credit)
            tb_data.append({
                'Account ID': acc_id,
                'Account Name': acc_name,
                'TB Debit': debit,
                'TB Credit': credit,
                'TB Net Balance': net_balance
            })
            
    return pd.DataFrame(tb_data)


def _clean_thai_pdf_text(text):
    """Clean Thai font encoding artifacts from PDF text.
    
    Some PDF fonts use private-use Unicode area (U+F700-U+F74F) for Thai characters.
    These look like garbled text but map to normal Thai characters.
    This function strips them back to readable Thai text using a mapping table.
    """
    # Map private-use Thai font characters to normal Unicode Thai
    thai_font_map = {
        '\uf700': '\u0e40', '\uf701': '\u0e41', '\uf702': '\u0e42', '\uf703': '\u0e43',
        '\uf704': '\u0e44', '\uf705': '\u0e48', '\uf706': '\u0e49', '\uf707': '\u0e4a',
        '\uf708': '\u0e4b', '\uf709': '\u0e4c', '\uf70a': '\u0e48', '\uf70b': '\u0e49',
        '\uf70c': '\u0e4a', '\uf70d': '\u0e4b', '\uf70e': '\u0e4c', '\uf70f': '\u0e4d',
        '\uf710': '\u0e31', '\uf711': '\u0e34', '\uf712': '\u0e35', '\uf713': '\u0e36',
        '\uf714': '\u0e37', '\uf715': '\u0e38', '\uf716': '\u0e39', '\uf717': '\u0e47',
        '\uf718': '\u0e48', '\uf719': '\u0e49', '\uf71a': '\u0e4a', '\uf71b': '\u0e4b',
        '\uf71c': '\u0e4c',
    }
    for char, replacement in thai_font_map.items():
        text = text.replace(char, replacement)
    # Remove any remaining private-use characters
    text = re.sub(r'[\uf700-\uf74f]', '', text)
    return text


def parse_tb_working_paper_pdf(pdf_file):
    """Parse กระดาษทำการ (Working Paper) in PDF format.
    
    This PDF has 3 column groups (each with Debit/Credit):
      1. งบทดลอง   (TB)  — The full trial balance (Debit = Credit for every account)
      2. งบกำไรขาดทุน (P&L) — P&L accounts only
      3. งบดุล       (BS)  — Balance sheet accounts only
    
    The net balance for each account is the LAST number printed on that line
    (either the BS or P&L column, whichever applies).
    We do NOT use the งบทดลอง pair because Debit always equals Credit there.
    """
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    
    lines = text.split('\n')
    tb_data = []
    
    for raw_line in lines:
        line = raw_line.rstrip('\r\n')
        stripped = line.strip()
        
        # Match lines starting with an account ID like '  1111-00'
        match = re.match(r'^\s*(\d{4}-\d{2})\s+(.*)', line)
        if not match:
            continue
        
        acc_id = match.group(1)
        rest = match.group(2)
        
        # Find ALL numbers on this line
        all_nums = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_nums:
            continue
        
        # The LAST number is the net balance (งบดุล or งบกำไรขาดทุน column)
        net_balance = float(all_nums[-1].replace(',', ''))
        
        # Derive a pseudo debit/credit: for BS accounts (1/2/3), treat as debit-normal;
        # for P&L (4/5), treat as debit-normal. We just store net balance.
        # For TB Debit/Credit columns, use the first number (งบทดลอง debit side)
        tb_gross = float(all_nums[0].replace(',', ''))
        
        # Extract account name: text between the account ID column and the first number
        name_match = re.match(r'^(.+?)(?=\s{3,}[\d,]+\.\d{2})', rest)
        if name_match:
            acc_name = _clean_thai_pdf_text(name_match.group(1).strip())
            if re.match(r'^[\d,]+\.\d{2}$', acc_name):
                acc_name = "Unknown"
        else:
            acc_name = "Unknown"
        
        tb_data.append({
            'Account ID': acc_id,
            'Account Name': acc_name,
            'TB Debit': net_balance,   # Net balance stored as debit for compatibility
            'TB Credit': 0.0,
            'TB Net Balance': net_balance
        })
    
    return pd.DataFrame(tb_data)

def parse_gl(pdf_file):
    # Read PDF text
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
        
    lines = text.split('\n')
    
    gl_data = []
    current_acc_id = None
    
    for i in range(len(lines)):
        line = lines[i].strip()
        
        # Check if line starts with an account ID (e.g., 1111-00)
        match_id = re.match(r'^(\d{4}-\d{2})\s+(.*)', line)
        if match_id:
            current_acc_id = match_id.group(1)
            rest_of_line = match_id.group(2)
            
            # Check if we already have it
            existing = next((item for item in gl_data if item["Account ID"] == current_acc_id), None)
            
            if not existing:
                # Extract account name from GL (first word/phrase before long spaces or parentheses)
                acc_name_match = re.match(r'^([^(\s]+(?:\s[^(\s]+)*)', rest_of_line)
                gl_acc_name = acc_name_match.group(1).strip() if acc_name_match else "Unknown"
                
                # First time seeing this account, try to extract BF
                bf_val = 0.0
                bf_match = re.search(r'\s+((?:\()?[\d,]+\.\d{2}(?:\))?)$', rest_of_line)
                if bf_match:
                    bf_str = bf_match.group(1)
                    is_negative = '(' in bf_str
                    bf_val = float(bf_str.replace('(', '').replace(')', '').replace(',', ''))
                    if is_negative:
                        bf_val = -bf_val
                        
                # Initialize net balance with BF for Balance Sheet accounts (1=Asset, 2=Liability, 3=Equity)
                if current_acc_id.startswith('1') or current_acc_id.startswith('2') or current_acc_id.startswith('3'):
                    initial_net = abs(bf_val)
                else:
                    initial_net = 0.0
                        
                gl_data.append({
                    'Account ID': current_acc_id,
                    'GL Account Name': gl_acc_name,
                    'GL Brought Forward': bf_val,
                    'GL Debit': 0.0,
                    'GL Credit': 0.0,
                    'GL Net Balance': initial_net
                })
            
        # Look for "รวม" followed by amounts
        match_total = re.search(r'รวม\s+([\d,.]+)\s+([\d,.]+)', line)
        if match_total and current_acc_id is not None:
            debit_str = match_total.group(1).replace(',', '')
            credit_str = match_total.group(2).replace(',', '')
            
            try:
                debit_val = float(debit_str)
            except ValueError:
                debit_val = 0.0
                
            try:
                credit_val = float(credit_str)
            except ValueError:
                credit_val = 0.0
                
            # Update the existing record
            existing = next((item for item in gl_data if item["Account ID"] == current_acc_id), None)
            if existing:
                existing['GL Debit'] = debit_val
                existing['GL Credit'] = credit_val
                
                # Apply BF logic for assets (1), liabilities (2), and equity (3)
                if current_acc_id.startswith('1') or current_acc_id.startswith('2') or current_acc_id.startswith('3'):
                    existing['GL Net Balance'] = abs(existing['GL Brought Forward'] + debit_val - credit_val)
                else:
                    existing['GL Net Balance'] = abs(debit_val - credit_val)
                    
    return pd.DataFrame(gl_data)

def main():
    st.title("GL & TB Reconciliation Tool")
    st.markdown("Upload your General Ledger (PDF), Trial Balance (Excel), and optionally Trial Balance (PDF) to verify Brought Forward balances.")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        tb_file = st.file_uploader("Upload Trial Balance / กระดาษทำการ (Excel or PDF)", type=["xls", "xlsx", "pdf"])
    with col2:
        gl_file = st.file_uploader("Upload General Ledger / บัญชีแยกประเภท (PDF)", type=["pdf"])
    with col3:
        tb_pdf_file = st.file_uploader("Upload Trial Balance for BF Check / งบทดลอง (PDF)", type=["pdf"])
        
    if st.button("Run Reconciliation", type="primary"):
        if not tb_file or not gl_file:
            st.error("Please upload both files first.")
            return
            
        with st.spinner("Processing files... This may take a moment."):
            # Parse TB (support both Excel and PDF กระดาษทำการ)
            try:
                tb_filename = tb_file.name.lower()
                if tb_filename.endswith('.pdf'):
                    tb_df = parse_tb_working_paper_pdf(tb_file)
                    st.info("ตรวจพบ กระดาษทำการ PDF — ใช้ตัวเลขคอลัมน์งบทดลอง (เดบิต/เครดิต) จากไฟล์นี้")
                else:
                    tb_df = parse_tb(tb_file)
            except Exception as e:
                st.error(f"Error reading TB file: {e}")
                return
                
            # Parse GL
            try:
                gl_df = parse_gl(gl_file)
            except Exception as e:
                st.error(f"Error reading GL PDF file: {e}")
                return
                
            # Reconciliation
            if tb_df.empty:
                st.error("Could not extract any accounts from TB.")
                return
            if gl_df.empty:
                st.error("Could not extract any accounts from GL.")
                return
                
            # Merge dataframes
            merged_df = pd.merge(tb_df, gl_df, on='Account ID', how='outer')
            
            # Parse and merge TB PDF if provided
            if tb_pdf_file:
                try:
                    tb_bf_df = parse_tb_pdf(tb_pdf_file)
                    merged_df = pd.merge(merged_df, tb_bf_df, on='Account ID', how='left')
                except Exception as e:
                    st.warning(f"Error reading TB PDF file: {e}. Skipping BF check.")
                    merged_df['TB Brought Forward Net'] = 0.0
            else:
                merged_df['TB Brought Forward Net'] = 0.0
            
            # Combine Account Names
            if 'Account Name' in merged_df.columns and 'GL Account Name' in merged_df.columns:
                merged_df['Account Name'] = merged_df['Account Name'].fillna(merged_df['GL Account Name'])
            elif 'GL Account Name' in merged_df.columns:
                merged_df['Account Name'] = merged_df['GL Account Name']
            
            # Identify missing status before fillna
            merged_df['Missing in TB original'] = pd.isna(merged_df['TB Net Balance'])
            merged_df['Missing in GL original'] = pd.isna(merged_df['GL Net Balance'])
            
            # Fill NaNs
            merged_df['TB Net Balance'] = merged_df['TB Net Balance'].fillna(0)
            merged_df['GL Net Balance'] = merged_df['GL Net Balance'].fillna(0)
            merged_df['TB Debit'] = merged_df['TB Debit'].fillna(0)
            merged_df['TB Credit'] = merged_df['TB Credit'].fillna(0)
            merged_df['GL Debit'] = merged_df['GL Debit'].fillna(0)
            merged_df['GL Credit'] = merged_df['GL Credit'].fillna(0)
            merged_df['TB Brought Forward Net'] = merged_df['TB Brought Forward Net'].fillna(0)
            merged_df['Account Name'] = merged_df['Account Name'].fillna("Unknown")
            
            # Calculate Differences
            # Floating point comparison using a small tolerance (e.g. 0.02)
            merged_df['Difference'] = abs(merged_df['GL Net Balance'] - merged_df['TB Net Balance'])
            merged_df['Is Match'] = merged_df['Difference'] <= 0.02
            
            merged_df['BF Difference'] = abs(abs(merged_df['GL Brought Forward']) - merged_df['TB Brought Forward Net'])
            
            # Status Logic
            def determine_status(row):
                if row['Missing in TB original']:
                    if row['GL Net Balance'] == 0:
                        return "Match (Zero Balance/Closed)"
                    return "Missing in TB"
                if row['Missing in GL original']:
                    if row['TB Net Balance'] == 0:
                        return "Match (Zero Balance/Closed)"
                    return "Missing in GL"
                if not row['Is Match']:
                    return "Amount Mismatch"
                return "Match"
                
            merged_df['Status'] = merged_df.apply(determine_status, axis=1)
            
            # Filter DataFrames
            mismatches_df = merged_df[merged_df['Status'] == 'Amount Mismatch']
            missing_in_gl_df = merged_df[merged_df['Status'] == 'Missing in GL']
            missing_in_tb_df = merged_df[merged_df['Status'] == 'Missing in TB']
            matches_df = merged_df[merged_df['Status'].str.contains('Match')]
            
            # Suspect Items (Top 10 largest absolute balance across TB and GL)
            merged_df['Max Absolute Balance'] = merged_df[['TB Net Balance', 'GL Net Balance']].max(axis=1)
            suspect_df = merged_df.sort_values(by='Max Absolute Balance', ascending=False).head(10)
            
            # Display Results
            st.success("Reconciliation Complete!")
            
            st.subheader("Summary")
            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
            col_s1.metric("Total Accounts", len(merged_df))
            col_s2.metric("Matched Accounts", len(matches_df))
            col_s3.metric("Mismatches", len(mismatches_df))
            col_s4.metric("Missing Accounts", len(missing_in_gl_df) + len(missing_in_tb_df))
            
            # Apply initial mapping
            if 'FS Line Item' not in merged_df.columns:
                merged_df['FS Line Item'] = merged_df.apply(auto_map, axis=1)
            
            tab1, tab2, tab3, tab4, tab_map, tab_export = st.tabs([
                "Discrepancies", "Missing Records", "Top 10 Suspects", "All Matches", "FS Mapping", "Export Reports"
            ])
            
            # Common numeric formatting configuration
            num_config = st.column_config.NumberColumn(format="%,.2f")
            common_col_config = {
                'GL Net Balance': st.column_config.NumberColumn("GL Net Balance (บัญชีแยกประเภท)", format="%,.2f"),
                'TB Net Balance': st.column_config.NumberColumn("TB Net Balance (กระดาษทำการ)", format="%,.2f"),
                'Difference': st.column_config.NumberColumn("Difference", format="%,.2f"),
                'GL Brought Forward': st.column_config.NumberColumn("GL Brought Forward (บัญชีแยกประเภท)", format="%,.2f"),
                'TB Brought Forward Net': st.column_config.NumberColumn("TB Brought Forward Net (งบทดลอง)", format="%,.2f"),
                'BF Difference': st.column_config.NumberColumn("BF Difference", format="%,.2f"),
                'Max Absolute Balance': st.column_config.NumberColumn("Max Absolute Balance", format="%,.2f"),
            }
            
            tab4_col_config = {
                'GL Net Balance': st.column_config.Column("GL Net Balance (บัญชีแยกประเภท)"),
                'TB Net Balance': st.column_config.Column("TB Net Balance (กระดาษทำการ)"),
                'GL Brought Forward': st.column_config.Column("GL Brought Forward (บัญชีแยกประเภท)"),
                'TB Brought Forward Net': st.column_config.Column("TB Brought Forward Net (งบทดลอง)"),
                'BF Difference': st.column_config.NumberColumn("BF Difference", format="%,.2f")
            }
            
            with tab1:
                st.write(f"Found {len(mismatches_df)} accounts with mismatched amounts.")
                st.dataframe(mismatches_df[['Account ID', 'Account Name', 'GL Net Balance', 'TB Net Balance', 'Difference', 'GL Brought Forward', 'TB Brought Forward Net', 'BF Difference', 'Status']], use_container_width=True, column_config=common_col_config)
                
            with tab2:
                st.write(f"Missing in GL: {len(missing_in_gl_df)}")
                st.dataframe(missing_in_gl_df[['Account ID', 'Account Name', 'TB Net Balance', 'Status']], use_container_width=True, column_config=common_col_config)
                st.write(f"Missing in TB: {len(missing_in_tb_df)}")
                st.dataframe(missing_in_tb_df[['Account ID', 'Account Name', 'GL Net Balance', 'Status']], use_container_width=True, column_config=common_col_config)
                
            with tab3:
                st.write("Top 10 largest accounts by absolute balance.")
                st.dataframe(suspect_df[['Account ID', 'Account Name', 'Max Absolute Balance', 'Status']], use_container_width=True, column_config=common_col_config)
                
            with tab4:
                st.write(f"Found {len(matches_df)} perfectly matched accounts.")
                
                display_matches_df = matches_df[['Account ID', 'Account Name', 'GL Net Balance', 'TB Net Balance', 'GL Brought Forward', 'TB Brought Forward Net', 'BF Difference', 'Status']].copy()
                
                # Format numeric columns to strings to preserve formatting when replacing with 'N/A'
                for col in ['GL Net Balance', 'TB Net Balance', 'GL Brought Forward', 'TB Brought Forward Net']:
                    display_matches_df[col] = display_matches_df[col].apply(lambda x: f"{x:,.2f}")
                
                # Replace with 'N/A' for missing accounts
                missing_tb_mask = matches_df['Missing in TB original']
                display_matches_df.loc[missing_tb_mask, 'TB Net Balance'] = 'N/A'
                display_matches_df.loc[missing_tb_mask, 'TB Brought Forward Net'] = 'N/A'
                
                missing_gl_mask = matches_df['Missing in GL original']
                display_matches_df.loc[missing_gl_mask, 'GL Net Balance'] = 'N/A'
                display_matches_df.loc[missing_gl_mask, 'GL Brought Forward'] = 'N/A'
                
                st.dataframe(display_matches_df, use_container_width=True, column_config=tab4_col_config)
                
            with tab_map:
                st.write("Review the auto-mapped Financial Statement Line Items. You can edit them directly in the table below.")
                st.info("The tool auto-guessed the line item based on Account ID and Name. Adjust any 'Unmapped' or incorrect items before exporting.")
                
                # Show editor
                mapping_df = merged_df[['Account ID', 'Account Name', 'TB Net Balance', 'FS Line Item']].copy()
                edited_mapping = st.data_editor(
                    mapping_df,
                    column_config={
                        "FS Line Item": st.column_config.SelectboxColumn(
                            "FS Line Item",
                            help="Select the FS Line Item for this account",
                            width="medium",
                            options=list(FS_LINE_ITEMS.keys()) + ["ไม่จัดประเภท (Unmapped)"],
                            required=True,
                        ),
                        "TB Net Balance": num_config
                    },
                    disabled=["Account ID", "Account Name", "TB Net Balance"],
                    use_container_width=True,
                    key="mapping_editor"
                )
                
                # Update merged_df with the edited mapping
                merged_df['FS Line Item'] = edited_mapping['FS Line Item']
                
                # Show live preview of the generated FS with inline breakdown
                st.subheader("Financial Statement Preview")
                st.write("Click on any line item to see the accounts that make up its total.")
                fs_preview = merged_df.groupby('FS Line Item')['TB Net Balance'].sum().reset_index()
                
                for _, row in fs_preview.iterrows():
                    fs_line = row['FS Line Item']
                    total = row['TB Net Balance']
                    if pd.notna(fs_line) and fs_line != "ไม่จัดประเภท (Unmapped)":
                        with st.expander(f"**{fs_line}** — Total: **{total:,.2f}**"):
                            line_items_df = merged_df[merged_df['FS Line Item'] == fs_line][['Account ID', 'Account Name', 'TB Net Balance']].reset_index(drop=True)
                            st.dataframe(line_items_df, use_container_width=True, column_config={'TB Net Balance': num_config})
                
                # Show unmapped items if any
                unmapped = merged_df[merged_df['FS Line Item'] == "ไม่จัดประเภท (Unmapped)"]
                if not unmapped.empty:
                    unmapped_total = unmapped['TB Net Balance'].sum()
                    with st.expander(f"⚠️ **ไม่จัดประเภท (Unmapped)** — Total: **{unmapped_total:,.2f}**"):
                        st.dataframe(unmapped[['Account ID', 'Account Name', 'TB Net Balance']].reset_index(drop=True), use_container_width=True, column_config={'TB Net Balance': num_config})


            with tab_export:
                st.write("Download your reconciled data or generate the finalized Financial Statements.")
                
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    # Original Reconciliation Export
                    buffer_recon = io.BytesIO()
                    with pd.ExcelWriter(buffer_recon, engine='xlsxwriter') as writer:
                        rename_dict = {
                            'TB Net Balance': 'TB Net Balance (กระดาษทำการ)',
                            'GL Net Balance': 'GL Net Balance (บัญชีแยกประเภท)',
                            'TB Brought Forward Net': 'TB Brought Forward Net (งบทดลอง)',
                            'GL Brought Forward': 'GL Brought Forward (บัญชีแยกประเภท)',
                            'TB Debit': 'TB Debit (กระดาษทำการ)',
                            'TB Credit': 'TB Credit (กระดาษทำการ)',
                            'GL Debit': 'GL Debit (บัญชีแยกประเภท)',
                            'GL Credit': 'GL Credit (บัญชีแยกประเภท)'
                        }
                        
                        export_df = merged_df.drop(columns=['Missing in TB original', 'Missing in GL original']).rename(columns=rename_dict)
                        export_df.to_excel(writer, sheet_name='All Accounts', index=False)
                        mismatches_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Discrepancies', index=False)
                        missing_in_tb_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Missing in TB', index=False)
                        missing_in_gl_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Missing in GL', index=False)
                        suspect_df.drop(columns=['Missing in TB original', 'Missing in GL original']).rename(columns=rename_dict).to_excel(writer, sheet_name='Top 10 Suspects', index=False)
                    
                    st.download_button(
                        label="Download Reconciliation Data",
                        data=buffer_recon.getvalue(),
                        file_name="Reconciliation_Report.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                
                with col_e2:
                    st.write("Export directly to the FS.xlsx template based on your mapping.")
                    template_path = "../FS.xlsx"
                    if os.path.exists(template_path):
                        # Group by mapping to get final sums
                        fs_summary_dict = merged_df.groupby('FS Line Item')['TB Net Balance'].sum().to_dict()
                        
                        try:
                            wb = openpyxl.load_workbook(template_path)
                            for fs_line, val in fs_summary_dict.items():
                                if fs_line in FS_CELLS:
                                    sheet_name = FS_CELLS[fs_line]["sheet"]
                                    cell_ref = FS_CELLS[fs_line]["cell"]
                                    wb[sheet_name][cell_ref].value = val
                            
                            buffer_fs = io.BytesIO()
                            wb.save(buffer_fs)
                            buffer_fs.seek(0)
                            
                            st.download_button(
                                label="Download Financial Statements (FS.xlsx)",
                                data=buffer_fs.getvalue(),
                                file_name="Generated_FS.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                        except Exception as e:
                            st.error(f"Error generating FS: {e}")
                    else:
                        st.error("Template 'FS.xlsx' not found in the application directory.")


if __name__ == "__main__":
    main()
