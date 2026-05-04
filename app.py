import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io
import math
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers as xl_numbers
from openpyxl.utils import get_column_letter
import os

# ==========================================
# 1. CONFIGURATION DICTIONARIES
# ==========================================
FS_LINE_ITEMS = {
    # Balance Sheet — Assets (use BS Debit column = งบดุล เดบิต)
    "เงินสดและรายการเทียบเท่าเงินสด": {"prefixes": ["111"], "keywords": ["เงินสด", "เงินฝาก"], "side": "bs_debit"},
    "ลูกหนี้การค้า": {"prefixes": ["113"], "keywords": ["ลูกหนี้"], "side": "bs_debit"},
    "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน": {"prefixes": ["121"], "keywords": ["เงินให้กู้ยืม"], "side": "bs_debit"},
    "สินทรัพย์หมุนเวียนอื่น": {"prefixes": ["115", "119", "150"], "keywords": ["ภาษีถูกหัก", "จ่ายล่วงหน้า", "ดอกเบี้ยค้างรับ"], "side": "bs_debit"},
    "อุปกรณ์-สุทธิ (Gross)": {"prefixes": ["141"], "keywords": ["เครื่องมือ", "เครื่องจักร", "อุปกรณ์สำนักงาน"], "side": "bs_debit"},
    "ค่าเสื่อมราคาสะสม": {"prefixes": ["142"], "keywords": ["ค่าเสื่อมราคาสะสม"], "side": "bs_credit"},
    
    # Balance Sheet — Liabilities (use BS Credit column = งบดุล เครดิต)
    "เจ้าหนี้การค้า": {"prefixes": ["212"], "keywords": ["เจ้าหนี้การค้า"], "side": "bs_credit"},
    "เจ้าหนี้หมุนเวียนอื่น": {"prefixes": ["211", "2131"], "keywords": ["เจ้าหนี้", "ค้างจ่าย", "สอบบัญชี", "ทำบัญชี"], "side": "bs_credit"},
    "หนี้สินหมุนเวียนอื่น": {"prefixes": ["2132", "2137", "2131-04"], "keywords": ["ภาษีหัก", "ภงด", "กรมสรรพากร", "ประกันสังคม", "รอนำส่ง"], "side": "bs_credit"},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"prefixes": ["2138"], "keywords": ["เงินกู้ยืม"], "side": "bs_credit"},
    
    # Balance Sheet — Equity (use BS Credit column = งบดุล เครดิต)
    "ทุนเรือนหุ้น": {"prefixes": ["31"], "keywords": ["ทุน"], "side": "bs_credit"},
    "กำไร(ขาดทุน)สะสม": {"prefixes": ["32"], "keywords": ["กำไร"], "side": "bs_debit"},
    
    # P&L — Revenue (use PL Credit column = งบกำไรขาดทุน เครดิต)
    "รายได้จากการให้บริการ": {"prefixes": ["41"], "keywords": ["รายได้จากการ"], "side": "pl_credit"},
    "รายได้อื่น": {"prefixes": ["42"], "keywords": ["รายได้อื่น", "ดอกเบี้ยรับ"], "side": "pl_credit"},
    
    # P&L — Expenses (use PL Debit column = งบกำไรขาดทุน เดบิต)
    "ภาษีเงินได้": {"prefixes": [], "keywords": ["ภาษีเงินได้นิติบุคคล"], "side": "pl_debit"}, # Added for Tax Engine
    "ต้นทุนการให้บริการ": {"prefixes": ["51"], "keywords": ["ต้นทุน", "ซื้อ"], "side": "pl_debit"},
    "ค่าใช้จ่ายในการขายและบริหาร": {"prefixes": ["52", "53"], "keywords": ["ค่าใช้จ่าย", "เงินเดือน", "ค่าธรรมเนียม", "ค่าเสื่อม", "ค่าเช่า", "ประกัน", "สอบบัญชี", "บริการ"], "side": "pl_debit"},
}

FS_BS_STRUCTURE = [
    ('header', 'สินทรัพย์', None, None),
    ('header', 'สินทรัพย์หมุนเวียน', None, None),
    ('item',   'เงินสดและรายการเทียบเท่าเงินสด', 4, 'เงินสดและรายการเทียบเท่าเงินสด'),
    ('item',   'ลูกหนี้การค้า', 5, 'ลูกหนี้การค้า'),
    ('item',   'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน', None, 'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน'),
    ('item',   'สินทรัพย์หมุนเวียนอื่น', 6, 'สินทรัพย์หมุนเวียนอื่น'),
    ('subtotal','รวมสินทรัพย์หมุนเวียน', None, ['เงินสดและรายการเทียบเท่าเงินสด', 'ลูกหนี้การค้า', 'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน', 'สินทรัพย์หมุนเวียนอื่น']),
    ('spacer',  None, None, None),
    ('header', 'สินทรัพย์ไม่หมุนเวียน', None, None),
    ('item',   'อุปกรณ์-สุทธิ (Gross)', None, 'อุปกรณ์-สุทธิ (Gross)'),
    ('item',   'ค่าเสื่อมราคาสะสม', None, 'ค่าเสื่อมราคาสะสม'),
    ('subtotal','อุปกรณ์-สุทธิ', 7, ['อุปกรณ์-สุทธิ (Gross)', '-ค่าเสื่อมราคาสะสม']),
    ('subtotal','รวมสินทรัพย์ไม่หมุนเวียน', None, ['อุปกรณ์-สุทธิ']),
    ('subtotal','รวมสินทรัพย์', None, ['รวมสินทรัพย์หมุนเวียน', 'รวมสินทรัพย์ไม่หมุนเวียน']),
    ('spacer',  None, None, None),
    ('header', 'หนี้สินและส่วนของเจ้าของ', None, None),
    ('header', 'หนี้สินหมุนเวียน', None, None),
    ('item',   'เจ้าหนี้การค้า', None, 'เจ้าหนี้การค้า'),
    ('item',   'เจ้าหนี้หมุนเวียนอื่น', 8, 'เจ้าหนี้หมุนเวียนอื่น'),
    ('item',   'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน', 9, 'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน'),
    ('item',   'หนี้สินหมุนเวียนอื่น', 10, 'หนี้สินหมุนเวียนอื่น'),
    ('subtotal','รวมหนี้สินหมุนเวียน', None, ['เจ้าหนี้การค้า', 'เจ้าหนี้หมุนเวียนอื่น', 'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน', 'หนี้สินหมุนเวียนอื่น']),
    ('subtotal','รวมหนี้สิน', None, ['รวมหนี้สินหมุนเวียน']),
    ('spacer',  None, None, None),
    ('header', 'ส่วนของเจ้าของ', None, None),
    ('item',   'ทุนเรือนหุ้น', None, 'ทุนเรือนหุ้น'),
    ('item',   'กำไร(ขาดทุน)สะสม', None, 'กำไร(ขาดทุน)สะสม'),
    ('subtotal','รวมส่วนของเจ้าของ', None, ['ทุนเรือนหุ้น', 'กำไร(ขาดทุน)สะสม']),
    ('subtotal','รวมหนี้สินและส่วนของเจ้าของ', None, ['รวมหนี้สิน', 'รวมส่วนของเจ้าของ']),
]

FS_PL_STRUCTURE = [
    ('header',  'รายได้', None, None),
    ('item',    'รายได้จากการให้บริการ', None, 'รายได้จากการให้บริการ'),
    ('item',    'รายได้อื่น', None, 'รายได้อื่น'),
    ('subtotal','รวมรายได้', None, ['รายได้จากการให้บริการ', 'รายได้อื่น']),
    ('spacer',  None, None, None),
    ('header',  'ค่าใช้จ่าย', None, None),
    ('item',    'ต้นทุนการให้บริการ', None, 'ต้นทุนการให้บริการ'),
    ('item',    'ค่าใช้จ่ายในการขายและบริหาร', None, 'ค่าใช้จ่ายในการขายและบริหาร'),
    ('subtotal','รวมค่าใช้จ่าย', None, ['ต้นทุนการให้บริการ', 'ค่าใช้จ่ายในการขายและบริหาร']),
    ('item',    'ภาษีเงินได้', None, 'ภาษีเงินได้'),
    ('subtotal','กำไร(ขาดทุน)สุทธิ', None, ['รวมรายได้', '-รวมค่าใช้จ่าย', '-ภาษีเงินได้']),
]

# ==========================================
# SANITIZE CONFIGURATION (PREVENT WHITESPACE BUGS)
# ==========================================
FS_LINE_ITEMS = {k.strip(): v for k, v in FS_LINE_ITEMS.items()}

cleaned_bs = []
for row in FS_BS_STRUCTURE:
    r_type, r_label, r_note, r_key = row
    clean_label = r_label.strip() if isinstance(r_label, str) else r_label
    if isinstance(r_key, list): clean_key = [k.strip() for k in r_key]
    elif isinstance(r_key, str): clean_key = r_key.strip()
    else: clean_key = r_key
    cleaned_bs.append((r_type, clean_label, r_note, clean_key))
FS_BS_STRUCTURE = cleaned_bs

cleaned_pl = []
for row in FS_PL_STRUCTURE:
    r_type, r_label, r_note, r_key = row
    clean_label = r_label.strip() if isinstance(r_label, str) else r_label
    if isinstance(r_key, list): clean_key = [k.strip() for k in r_key]
    elif isinstance(r_key, str): clean_key = r_key.strip()
    else: clean_key = r_key
    cleaned_pl.append((r_type, clean_label, r_note, clean_key))
FS_PL_STRUCTURE = cleaned_pl


# ==========================================
# EXCEL GENERATION FUNCTIONS
# ==========================================
def build_fs_from_mapping(fs_mapping: dict, structure: list) -> dict:
    computed = {}
    for row_type, label, note, key in structure:
        if row_type == 'item':
            computed[label] = fs_mapping.get(key, 0.0) if key else 0.0
        elif row_type == 'subtotal' and isinstance(key, list):
            total = 0.0
            for k in key:
                if k.startswith('-'):
                    total -= computed.get(k[1:], 0.0)
                else:
                    total += computed.get(k, 0.0)
            computed[label] = total
        else:
            computed[label] = None
    return computed

def _apply_cell_style(cell, bold=False, italic=False, indent=0, bg=None, border_top=False, border_bottom=False, number_format=None, align='left'):
    cell.font = Font(name='Cordia New', size=14, bold=bold, italic=italic)
    cell.alignment = Alignment(horizontal=align, vertical='center', indent=indent, wrap_text=False)
    if bg:
        cell.fill = PatternFill('solid', fgColor=bg)
    if number_format:
        cell.number_format = number_format
    thin = Side(style='thin')
    double = Side(style='double')
    cell.border = Border(top=double if border_top else None, bottom=double if border_bottom else None)

def _write_fs_sheet(ws, structure, years_computed, year_labels, company_name, statement_title, period_label):
    NUM_FMT = '#,##0.00;[Red]-#,##0.00'
    HDR_FILL = 'DDEEFF'
    SUBTOT_FILL = 'F0F4F8'

    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 10
    for i, _ in enumerate(year_labels):
        col = get_column_letter(3 + i * 2) 
        ws.column_dimensions[col].width = 18
        ws.column_dimensions[get_column_letter(4 + i * 2)].width = 2

    row = 1
    ws.row_dimensions[row].height = 20
    ws.cell(row, 1, company_name)
    _apply_cell_style(ws.cell(row, 1), bold=True)
    row += 1
    ws.cell(row, 1, statement_title)
    _apply_cell_style(ws.cell(row, 1), bold=True)
    row += 1
    ws.cell(row, 1, period_label)
    _apply_cell_style(ws.cell(row, 1))
    row += 1
    ws.cell(row, 1, 'หน่วย : บาท')
    _apply_cell_style(ws.cell(row, 1), italic=True)
    row += 1

    ws.row_dimensions[row].height = 18
    for i, lbl in enumerate(year_labels):
        col = 3 + i * 2
        c = ws.cell(row, col, lbl)
        _apply_cell_style(c, bold=True, align='center')
    row += 1

    for row_type, label, note, key in structure:
        ws.row_dimensions[row].height = 18
        if row_type == 'spacer':
            row += 1
            continue

        is_header = row_type == 'header'
        is_subtotal = row_type == 'subtotal'
        indent = 0 if is_header else (1 if is_subtotal else 2)

        c = ws.cell(row, 1, label)
        _apply_cell_style(c, bold=(is_header or is_subtotal), indent=indent, bg=HDR_FILL if is_header else (SUBTOT_FILL if is_subtotal else None))

        if note:
            nc = ws.cell(row, 2, note)
            _apply_cell_style(nc, align='center')

        for i, lbl in enumerate(year_labels):
            col = 3 + i * 2
            val = years_computed[lbl].get(label) if label else None
            vc = ws.cell(row, col, val if val is not None else None)
            _apply_cell_style(vc, bold=is_subtotal, bg=SUBTOT_FILL if is_subtotal else None, number_format=NUM_FMT, align='right',
                              border_top=is_subtotal, border_bottom=(is_subtotal and 'รวม' in (label or '') and 'ทั้งหมด' not in (label or '')))
        row += 1

    ws.row_dimensions[row].height = 14
    ws.cell(row, 1, 'หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้')
    _apply_cell_style(ws.cell(row, 1), italic=True)


def generate_fs_excel(years_data: dict, company_name: str, current_year: str, prior_year: str = None, 
                      registered_capital: float = 0, shares: int = 0, par_value: float = 0) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    year_labels = [current_year]
    if prior_year and prior_year in years_data:
        year_labels.append(prior_year)

    bs_computed = {} 
    pl_computed = {} 
    for yl in year_labels:
        mapping = years_data.get(yl, {})
        pl_vals = build_fs_from_mapping(mapping, FS_PL_STRUCTURE)
        net_profit = pl_vals.get('กำไร(ขาดทุน)สุทธิ', 0.0) or 0.0
        mapping_bs = dict(mapping)
        mapping_bs['กำไร(ขาดทุน)สะสม'] = mapping.get('กำไร(ขาดทุน)สะสม', 0.0) + net_profit
        
        # Inject user's paid-up capital
        if registered_capital > 0:
            mapping_bs['ทุนเรือนหุ้น'] = registered_capital
            
        bs_computed[yl] = build_fs_from_mapping(mapping_bs, FS_BS_STRUCTURE)
        pl_computed[yl] = pl_vals

    ws_bs = wb.create_sheet('งบฐานะการเงิน')
    _write_fs_sheet(ws_bs, FS_BS_STRUCTURE, bs_computed, year_labels, company_name, 'งบฐานะการเงิน', f'ณ วันที่ 31 ธันวาคม {current_year}')

    ws_pl = wb.create_sheet('งบกำไรขาดทุน')
    _write_fs_sheet(ws_pl, FS_PL_STRUCTURE, pl_computed, year_labels, company_name, 'งบกำไรขาดทุน', f'สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม {current_year}')

    _write_equity_sheet(wb, company_name, current_year, prior_year, bs_computed, pl_computed, year_labels, registered_capital, shares, par_value)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _write_equity_sheet(wb, company_name, current_year, prior_year, bs_computed, pl_computed, year_labels, registered_capital, shares, par_value):
    ws = wb.create_sheet('งบส่วนของเจ้าของ')
    ws.column_dimensions['A'].width = 46
    for col in ['B', 'C', 'E', 'G']: ws.column_dimensions[col].width = 2
    for col in ['D', 'F', 'H']: ws.column_dimensions[col].width = 18
    NUM_FMT = '#,##0.00;[Red]-#,##0.00'

    def wc(r, c, v=None, bold=False, align='left', num=False, italic=False):
        cell = ws.cell(r, c, v)
        cell.font = Font(name='Cordia New', size=14, bold=bold, italic=italic)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        if num and v is not None:
            cell.number_format = NUM_FMT
        return cell

    r = 1
    wc(r, 1, company_name, bold=True); r += 1
    wc(r, 1, 'งบการเปลี่ยนแปลงส่วนของเจ้าของ', bold=True); r += 1
    wc(r, 1, f'สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม {current_year}'); r += 1
    wc(r, 1, 'หน่วย : บาท', italic=True); r += 1
    r += 1

    wc(r, 4, 'ทุนที่ออกและ', bold=True, align='center'); wc(r, 6, 'กำไร(ขาดทุน)', bold=True, align='center'); wc(r, 8, 'รวมส่วนของ', bold=True, align='center'); r += 1
    wc(r, 4, 'เรียกชำระแล้ว', bold=True, align='center'); wc(r, 6, 'สะสม', bold=True, align='center'); wc(r, 8, 'เจ้าของ', bold=True, align='center'); r += 1

    def get_pl(yl, key): return pl_computed.get(yl, {}).get(key, 0.0) or 0.0
    def get_bs(yl, key): return bs_computed.get(yl, {}).get(key, 0.0) or 0.0

    if prior_year and prior_year in bs_computed:
        prior_capital = get_bs(prior_year, 'ทุนเรือนหุ้น') or registered_capital
        prior_retained = get_bs(prior_year, 'กำไร(ขาดทุน)สะสม') - get_pl(prior_year, 'กำไร(ขาดทุน)สุทธิ')
        prior_net = get_pl(prior_year, 'กำไร(ขาดทุน)สุทธิ')

        wc(r, 1, f'ยอด ณ วันต้นปี {prior_year}', bold=True)
        wc(r, 4, prior_capital, num=True, align='right')
        wc(r, 6, prior_retained, num=True, align='right')
        wc(r, 8, prior_capital + prior_retained, num=True, align='right'); r += 1
        wc(r, 1, 'กำไร(ขาดทุน)สุทธิสำหรับปี')
        wc(r, 6, prior_net, num=True, align='right')
        wc(r, 8, prior_net, num=True, align='right'); r += 1
        prior_end_cap = prior_capital
        prior_end_ret = prior_retained + prior_net
        wc(r, 1, f'ยอดคงเหลือ ณ สิ้นปี {prior_year}', bold=True)
        wc(r, 4, prior_end_cap, num=True, align='right')
        wc(r, 6, prior_end_ret, num=True, align='right')
        wc(r, 8, prior_end_cap + prior_end_ret, num=True, align='right'); r += 1
        r += 1
        cur_capital_open = prior_end_cap
        cur_retained_open = prior_end_ret
    else:
        cur_capital_open = registered_capital
        cur_retained_open = 0

    cur_net = get_pl(current_year, 'กำไร(ขาดทุน)สุทธิ')
    cur_capital = get_bs(current_year, 'ทุนเรือนหุ้น') or registered_capital

    wc(r, 1, f'ยอด ณ วันต้นปี {current_year}', bold=True)
    wc(r, 4, cur_capital_open, num=True, align='right')
    wc(r, 6, cur_retained_open, num=True, align='right')
    wc(r, 8, cur_capital_open + cur_retained_open, num=True, align='right'); r += 1
    wc(r, 1, 'กำไร(ขาดทุน)สุทธิสำหรับปี')
    wc(r, 6, cur_net, num=True, align='right')
    wc(r, 8, cur_net, num=True, align='right'); r += 1
    cur_end_ret = cur_retained_open + cur_net
    wc(r, 1, f'ยอดคงเหลือ ณ สิ้นปี {current_year}', bold=True)
    wc(r, 4, cur_capital, num=True, align='right')
    wc(r, 6, cur_end_ret, num=True, align='right')
    wc(r, 8, cur_capital + cur_end_ret, num=True, align='right'); r += 1
    r += 1
    wc(r, 1, 'หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้', italic=True)


# ==========================================
# PARSERS
# ==========================================
def auto_map(row):
    acc_id = str(row.get('Account ID', '')).strip()
    acc_name = str(row.get('Account Name', '')).strip()
    
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
    for page in doc: text += page.get_text()
    lines = text.split('\n')
    tb_bf_data = []
    for line in lines:
        line = line.strip()
        match_id = re.match(r'^(\d{4}-\d{2})\s+(.*)', line)
        if match_id:
            acc_id = match_id.group(1).strip()
            rest_of_line = match_id.group(2)
            numbers = re.findall(r'((?:\()?[\d,]+\.\d{2}(?:\))?)', rest_of_line)
            if len(numbers) >= 6:
                try:
                    bf_dr = float(numbers[-6].replace(',', '').replace('(', '-').replace(')', ''))
                    bf_cr = float(numbers[-5].replace(',', '').replace('(', '-').replace(')', ''))
                    tb_bf_data.append({'Account ID': acc_id, 'TB Brought Forward Net': abs(bf_dr - bf_cr)})
                except: pass
    return pd.DataFrame(tb_bf_data)


st.set_page_config(page_title="Finance Reconcile", layout="wide", page_icon="📊")

def parse_tb(excel_file):
    df = pd.read_excel(excel_file, header=None)
    tb_data = []
    start_reading = False
    for index, row in df.iterrows():
        acc_id = str(row[0]).strip()
        if not start_reading:
            if re.match(r'^\d{4}-\d{2}$', acc_id): start_reading = True
            else: continue
                
        if start_reading:
            if pd.isna(row[0]) or str(row[0]).strip() == 'nan': continue
            acc_id = str(row[0]).strip()
            if not re.match(r'^\d{4}-\d{2}$', acc_id): continue
            acc_name = str(row[1]).strip()
            
            debit = float(row[2]) if not pd.isna(row[2]) else 0.0
            credit = float(row[3]) if not pd.isna(row[3]) else 0.0
            pl_debit = float(row[4]) if len(row) > 4 and not pd.isna(row[4]) else 0.0
            pl_credit = float(row[5]) if len(row) > 5 and not pd.isna(row[5]) else 0.0
            bs_debit = float(row[6]) if len(row) > 6 and not pd.isna(row[6]) else 0.0
            bs_credit = float(row[7]) if len(row) > 7 and not pd.isna(row[7]) else 0.0
            
            tb_data.append({
                'Account ID': acc_id, 'Account Name': acc_name,
                'TB Debit': debit, 'TB Credit': credit, 'TB Net Balance': abs(debit - credit),
                'PL Debit': pl_debit, 'PL Credit': pl_credit, 'BS Debit': bs_debit, 'BS Credit': bs_credit,
            })
    return pd.DataFrame(tb_data)

def _clean_thai_pdf_text(text):
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
    for char, replacement in thai_font_map.items(): text = text.replace(char, replacement)
    text = re.sub(r'[\uf700-\uf74f]', '', text)
    return text

def parse_tb_working_paper_pdf(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc: text += page.get_text()
    lines = text.split('\n')
    tb_data = []
    
    for raw_line in lines:
        line = raw_line.rstrip('\r\n')
        match = re.match(r'^\s*(\d{4}-\d{2})\s+(.*)', line)
        if not match: continue
        
        acc_id = match.group(1).strip()
        rest = match.group(2)
        all_nums = re.findall(r'([\d,]+\.\d{2})', rest)
        if not all_nums: continue
        
        name_match = re.match(r'^(.+?)(?=\s{3,}[\d,]+\.\d{2})', rest)
        acc_name = _clean_thai_pdf_text(name_match.group(1).strip()) if name_match else "Unknown"
        if re.match(r'^[\d,]+\.\d{2}$', acc_name): acc_name = "Unknown"
        
        is_bs_account = acc_id[0] in ('1', '2', '3')
        is_debit_normal = acc_id[0] in ('1', '5')
        
        tb_debit = tb_credit = pl_debit = pl_credit = bs_debit = bs_credit = 0.0
        
        if len(all_nums) >= 1:
            val = float(all_nums[0].replace(',', ''))
            if is_debit_normal: tb_debit = val
            else: tb_credit = val
        
        if len(all_nums) >= 2:
            last_val = float(all_nums[-1].replace(',', ''))
            if is_bs_account:
                if is_debit_normal: bs_debit = last_val
                else: bs_credit = last_val
            else:
                if acc_id[0] == '4': pl_credit = last_val
                else: pl_debit = last_val
        elif len(all_nums) == 1:
            val = float(all_nums[0].replace(',', ''))
            if is_bs_account:
                if is_debit_normal: bs_debit = val
                else: bs_credit = val
            else:
                if acc_id[0] == '4': pl_credit = val
                else: pl_debit = val
        
        tb_data.append({
            'Account ID': acc_id, 'Account Name': acc_name,
            'TB Debit': tb_debit, 'TB Credit': tb_credit, 'TB Net Balance': abs(tb_debit - tb_credit),
            'PL Debit': pl_debit, 'PL Credit': pl_credit, 'BS Debit': bs_debit, 'BS Credit': bs_credit,
        })
    return pd.DataFrame(tb_data)

def parse_gl(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc: text += page.get_text()
    lines = text.split('\n')
    
    extracted_company = ""
    extracted_year = ""
    for i in range(min(10, len(lines))):
        line = lines[i].strip()
        if "บริษัท" in line and not extracted_company:
            match_comp = re.split(r'\s{3,}|\t|หน้า|หนา', line)
            extracted_company = match_comp[0].strip()
        if "วันที่จาก" in line and not extracted_year:
            match_yr = re.search(r'(25\d{2})', line)
            if match_yr: extracted_year = "พ.ศ. " + match_yr.group(1)
                
    gl_data = []
    current_acc_id = None
    
    for i in range(len(lines)):
        line = lines[i].strip()
        match_id = re.match(r'^(\d{4}-\d{2})\s+(.*)', line)
        if match_id:
            current_acc_id = match_id.group(1).strip()
            rest_of_line = match_id.group(2)
            existing = next((item for item in gl_data if item["Account ID"] == current_acc_id), None)
            
            if not existing:
                acc_name_match = re.match(r'^([^(\s]+(?:\s[^(\s]+)*)', rest_of_line)
                gl_acc_name = acc_name_match.group(1).strip() if acc_name_match else "Unknown"
                
                bf_val = 0.0
                bf_match = re.search(r'\s+((?:\()?[\d,]+\.\d{2}(?:\))?)$', rest_of_line)
                if bf_match:
                    bf_str = bf_match.group(1)
                    bf_val = float(bf_str.replace('(', '').replace(')', '').replace(',', ''))
                    if '(' in bf_str: bf_val = -bf_val
                        
                initial_net = abs(bf_val) if current_acc_id[0] in ['1', '2', '3'] else 0.0
                gl_data.append({
                    'Account ID': current_acc_id, 'GL Account Name': gl_acc_name,
                    'GL Brought Forward': bf_val, 'GL Debit': 0.0, 'GL Credit': 0.0, 'GL Net Balance': initial_net
                })
            
        match_total = re.search(r'รวม\s+([\d,.]+)\s+([\d,.]+)', line)
        if match_total and current_acc_id is not None:
            debit_val = float(match_total.group(1).replace(',', '')) if match_total.group(1) else 0.0
            credit_val = float(match_total.group(2).replace(',', '')) if match_total.group(2) else 0.0
                
            existing = next((item for item in gl_data if item["Account ID"] == current_acc_id), None)
            if existing:
                existing['GL Debit'] = debit_val
                existing['GL Credit'] = credit_val
                if current_acc_id[0] in ['1', '2', '3']:
                    existing['GL Net Balance'] = abs(existing['GL Brought Forward'] + debit_val - credit_val)
                else:
                    existing['GL Net Balance'] = abs(debit_val - credit_val)
                    
    return pd.DataFrame(gl_data), extracted_company, extracted_year

# ==========================================
# MAIN APP 
# ==========================================
def main():
    st.title("GL & TB Reconciliation Tool")
    st.markdown("Upload your General Ledger (PDF), Trial Balance (Excel), and optionally Trial Balance (PDF) to verify Brought Forward balances.")
    
    default_company = st.session_state.get('_gl_extracted_company', '')
    default_year = st.session_state.get('_gl_extracted_year', 'พ.ศ. 2568')
    default_prior_year = st.session_state.get('_gl_extracted_prior_year', 'พ.ศ. 2567')
    
    # ---------------------------------------------
    # UPDATED COMPANY INFO SECTION (Authorized vs Paid-up)
    # ---------------------------------------------
    with st.expander("ข้อมูลบริษัท / Company Info (สำหรับสร้างงบการเงิน)", expanded=True):
        col_ci1, col_ci2 = st.columns([3, 1])
        with col_ci1:
            company_name = st.text_input("ชื่อบริษัท", value=default_company, placeholder="บริษัท xxxxxxx จำกัด", key='company_name')
        with col_ci2:
            current_year_label = st.text_input("ปีปัจจุบัน", value=default_year, key='current_year_label')
        
        st.write("---")
        st.write("##### 📊 ข้อมูลทุนเรือนหุ้น (Share Capital)")
        st.caption("ข้อมูลนี้จะถูกนำไปจัดหน้าในส่วนของ 'ส่วนของเจ้าของ' ในงบฐานะการเงิน")
        
        col_cap1, col_cap2 = st.columns(2)
        with col_cap1:
            st.markdown("**1. ทุนจดทะเบียน (Authorized Capital)**")
            reg_shares = st.number_input("จำนวนหุ้นจดทะเบียน (หุ้น)", min_value=0, value=10000, step=1000, key='reg_shares')
            reg_par = st.number_input("มูลค่าจดทะเบียนหุ้นละ (บาท)", min_value=0.0, value=100.0, step=1.0, key='reg_par')
            st.info(f"💡 ทุนจดทะเบียนรวม: **{reg_shares * reg_par:,.2f}** บาท\n\n*(แสดงเป็นข้อมูลในงบดุล แต่ไม่นำไปบวกรวมในยอดหนี้สินและส่วนของเจ้าของ)*")
            
        with col_cap2:
            st.markdown("**2. ทุนที่ออกและเรียกชำระแล้ว (Paid-up Capital)**")
            paid_shares = st.number_input("จำนวนหุ้นที่เรียกชำระ (หุ้น)", min_value=0, value=10000, step=1000, key='paid_shares')
            paid_par = st.number_input("มูลค่าที่เรียกชำระแล้วหุ้นละ (บาท)", min_value=0.0, value=25.0, step=1.0, key='paid_par')
            st.success(f"💡 ทุนที่เรียกชำระแล้วรวม: **{paid_shares * paid_par:,.2f}** บาท\n\n*(ยอดนี้คือตัวเลขจริงที่จะถูกนำไปคำนวณเพื่อให้งบดุลลงตัว)*")
        
        st.write("---")
        st.selectbox("วิธีการคำนวณภาษีเงินได้นิติบุคคล (Corporate Tax Rule)", 
                     ["Auto-Detect จากงบทดลอง (แนะนำ)", "SME (ยกเว้น 300k แรก, 15%-20%)", "Standard (20%)", "ไม่คำนวณอัตโนมัติ (Manual)"], 
                     key='tax_method')
    
    col1, col2, col3 = st.columns(3)
    with col1: tb_file = st.file_uploader("Upload Trial Balance / กระดาษทำการ (Excel or PDF)", type=["xls", "xlsx", "pdf"])
    with col2: gl_file = st.file_uploader("Upload General Ledger / บัญชีแยกประเภท (PDF)", type=["pdf"])
    with col3: tb_pdf_file = st.file_uploader("Upload Trial Balance for BF Check / งบทดลอง (PDF)", type=["pdf"])
    
    with st.expander("📊 อัปโหลดกระดาษทำการปีก่อนหน้า (สำหรับสร้างงบเปรียบเทียบ)", expanded=False):
        col_py1, col_py2 = st.columns([1, 2])
        with col_py1: prior_year_label = st.text_input("ปีก่อนหน้า", value=default_prior_year, key='prior_year_label')
        with col_py2: prior_tb_file = st.file_uploader("Upload Prior Year กระดาษทำการ (Excel or PDF)", type=["xls", "xlsx", "pdf"], key='prior_tb_upload')
        
    run_clicked = st.button("Run Reconciliation", type="primary")
    should_process = False
    if run_clicked:
        if not tb_file or not gl_file:
            st.error("Please upload both files first.")
            return
        should_process = True

    if should_process:
        with st.spinner("Processing files... This may take a moment."):
            try:
                tb_filename = tb_file.name.lower()
                if tb_filename.endswith('.pdf'):
                    tb_df = parse_tb_working_paper_pdf(tb_file)
                    st.info("ตรวจพบ กระดาษทำการ PDF — ใช้ตัวเลขคอลัมน์งบทดลอง (เดบิต/เครดิต) จากไฟล์นี้")
                else: tb_df = parse_tb(tb_file)
            except Exception as e:
                st.error(f"Error reading TB file: {e}")
                return
                
            try:
                gl_df, extracted_company, extracted_year = parse_gl(gl_file)
                if extracted_company: st.session_state['_gl_extracted_company'] = extracted_company; company_name = company_name or extracted_company
                if extracted_year:
                    st.session_state['_gl_extracted_year'] = extracted_year
                    if current_year_label == 'พ.ศ. 2568' or not current_year_label: current_year_label = extracted_year
                    try:
                        # Pull the regex search outside the f-string to avoid the backslash error
                        year_match = re.search(r'\d{4}', extracted_year)
                        prior_yr = f"พ.ศ. {int(year_match.group()) - 1}"
                        
                        st.session_state['_gl_extracted_prior_year'] = prior_yr
                        if prior_year_label == 'พ.ศ. 2567' or not prior_year_label: 
                            prior_year_label = prior_yr
                    except: 
                        pass
                st.info(f"📋 ข้อมูลจาก GL: **{extracted_company}** | ปี: **{extracted_year}**")
            except Exception as e:
                st.error(f"Error reading GL PDF file: {e}")
                return
                
            if tb_df.empty or gl_df.empty:
                st.error("Could not extract accounts from TB or GL.")
                return
                
            merged_df = pd.merge(tb_df, gl_df, on='Account ID', how='outer')
            
            if tb_pdf_file:
                try: merged_df = pd.merge(merged_df, parse_tb_pdf(tb_pdf_file), on='Account ID', how='left')
                except: merged_df['TB Brought Forward Net'] = 0.0
            else: merged_df['TB Brought Forward Net'] = 0.0
            
            if 'Account Name' in merged_df.columns and 'GL Account Name' in merged_df.columns:
                merged_df['Account Name'] = merged_df['Account Name'].fillna(merged_df['GL Account Name'])
            elif 'GL Account Name' in merged_df.columns:
                merged_df['Account Name'] = merged_df['GL Account Name']
            
            merged_df['Missing in TB original'] = pd.isna(merged_df['TB Net Balance'])
            merged_df['Missing in GL original'] = pd.isna(merged_df['GL Net Balance'])
            
            # Clean spaces from IDs and Names
            merged_df['Account ID'] = merged_df['Account ID'].str.strip()
            merged_df['Account Name'] = merged_df['Account Name'].str.strip().fillna("Unknown")

            for col in ['TB Net Balance', 'GL Net Balance', 'TB Debit', 'TB Credit', 'GL Debit', 'GL Credit', 'TB Brought Forward Net', 'PL Debit', 'PL Credit', 'BS Debit', 'BS Credit']:
                if col not in merged_df.columns: merged_df[col] = 0.0
                merged_df[col] = merged_df[col].fillna(0)
            
            merged_df['Difference'] = abs(merged_df['GL Net Balance'] - merged_df['TB Net Balance'])
            merged_df['Is Match'] = merged_df['Difference'] <= 0.02
            merged_df['BF Difference'] = abs(abs(merged_df['GL Brought Forward']) - merged_df['TB Brought Forward Net'])
            
            def determine_status(row):
                if row['Missing in TB original']: return "Match (Zero Balance/Closed)" if row['GL Net Balance'] == 0 else "Missing in TB"
                if row['Missing in GL original']: return "Match (Zero Balance/Closed)" if row['TB Net Balance'] == 0 else "Missing in GL"
                if not row['Is Match']: return "Amount Mismatch"
                return "Match"
                
            merged_df['Status'] = merged_df.apply(determine_status, axis=1)
            
            mismatches_df = merged_df[merged_df['Status'] == 'Amount Mismatch']
            missing_in_gl_df = merged_df[merged_df['Status'] == 'Missing in GL']
            missing_in_tb_df = merged_df[merged_df['Status'] == 'Missing in TB']
            matches_df = merged_df[merged_df['Status'].str.contains('Match')]
            
            merged_df['Max Absolute Balance'] = merged_df[['TB Net Balance', 'GL Net Balance']].max(axis=1)
            suspect_df = merged_df.sort_values(by='Max Absolute Balance', ascending=False).head(10)
            
            if 'FS Line Item' not in merged_df.columns:
                merged_df['FS Line Item'] = merged_df.apply(auto_map, axis=1)
            st.session_state['recon_merged_df'] = merged_df
            st.session_state['recon_mismatches_df'] = mismatches_df
            st.session_state['recon_missing_gl_df'] = missing_in_gl_df
            st.session_state['recon_missing_tb_df'] = missing_in_tb_df
            st.session_state['recon_matches_df'] = matches_df
            st.session_state['recon_suspect_df'] = suspect_df
            st.session_state['data_parsed'] = True

    if st.session_state.get('data_parsed', False):
        merged_df = st.session_state['recon_merged_df']
        mismatches_df = st.session_state['recon_mismatches_df']
        missing_in_gl_df = st.session_state['recon_missing_gl_df']
        missing_in_tb_df = st.session_state['recon_missing_tb_df']
        matches_df = st.session_state['recon_matches_df']
        suspect_df = st.session_state['recon_suspect_df']

        st.success("Reconciliation Complete!")
        st.subheader("Summary")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Total Accounts", len(merged_df))
        col_s2.metric("Matched Accounts", len(matches_df))
        col_s3.metric("Mismatches", len(mismatches_df))
        col_s4.metric("Missing Accounts", len(missing_in_gl_df) + len(missing_in_tb_df))

        tab1, tab2, tab3, tab4, tab_map, tab_export = st.tabs([
            "Discrepancies", "Missing Records", "Top 10 Suspects", "All Matches", "FS Mapping", "Export Reports"
        ])

        common_col_config = {
            'GL Net Balance': st.column_config.NumberColumn("GL Net", format="%,.2f"),
            'TB Net Balance': st.column_config.NumberColumn("TB Net", format="%,.2f"),
            'Difference': st.column_config.NumberColumn("Diff", format="%,.2f"),
            'GL Brought Forward': st.column_config.NumberColumn("GL BF", format="%,.2f"),
            'TB Brought Forward Net': st.column_config.NumberColumn("TB BF", format="%,.2f"),
            'BF Difference': st.column_config.NumberColumn("BF Diff", format="%,.2f"),
            'Max Absolute Balance': st.column_config.NumberColumn("Max Balance", format="%,.2f"),
        }

        with tab1: st.dataframe(mismatches_df[['Account ID', 'Account Name', 'GL Net Balance', 'TB Net Balance', 'Difference', 'GL Brought Forward', 'TB Brought Forward Net', 'BF Difference', 'Status']], use_container_width=True, column_config=common_col_config)
        with tab2: st.dataframe(missing_in_gl_df[['Account ID', 'Account Name', 'TB Net Balance', 'Status']], use_container_width=True, column_config=common_col_config); st.dataframe(missing_in_tb_df[['Account ID', 'Account Name', 'GL Net Balance', 'Status']], use_container_width=True, column_config=common_col_config)
        with tab3: st.dataframe(suspect_df[['Account ID', 'Account Name', 'Max Absolute Balance', 'Status']], use_container_width=True, column_config=common_col_config)
        with tab4: st.dataframe(matches_df[['Account ID', 'Account Name', 'GL Net Balance', 'TB Net Balance', 'GL Brought Forward', 'TB Brought Forward Net', 'BF Difference', 'Status']], use_container_width=True, column_config=common_col_config)

        # ---------------------------------------------
        # UPDATED TAB MAP (Includes Tax Engine & Subtotals)
        # ---------------------------------------------
        with tab_map:
            st.write("### 🗂️ Master FS Mapping & Verification")
            st.write("Review your mapping and see exactly where the data comes from before generating **FS.xlsx**.")

            # AUTO-DETECT & DYNAMIC CORPORATE TAX CALCULATION
            rev_total = merged_df[merged_df['Account ID'].str.startswith('4', na=False)]['PL Credit'].sum() - merged_df[merged_df['Account ID'].str.startswith('4', na=False)]['PL Debit'].sum()
            exp_total = merged_df[merged_df['Account ID'].str.startswith('5', na=False)]['PL Debit'].sum() - merged_df[merged_df['Account ID'].str.startswith('5', na=False)]['PL Credit'].sum()
            if rev_total == 0 and exp_total == 0:
                rev_total = merged_df[merged_df['Account ID'].str.startswith('4', na=False)]['TB Net Balance'].sum()
                exp_total = merged_df[merged_df['Account ID'].str.startswith('5', na=False)]['TB Net Balance'].sum()

            pre_tax_profit = rev_total - exp_total
            corporate_tax = 0.0
            
            tax_method = st.session_state.get('tax_method', "Auto-Detect จากงบทดลอง (แนะนำ)")
            applied_tax_rule = tax_method 

            if tax_method == "Auto-Detect จากงบทดลอง (แนะนำ)":
                has_tax_expense = merged_df['Account Name'].str.contains('ภาษีเงินได้นิติบุคคล|ค่าใช้จ่ายภาษีเงินได้', na=False).any()
                total_revenue = merged_df[merged_df['Account ID'].str.startswith('4', na=False)]['TB Net Balance'].sum()
                capital_df = merged_df[(merged_df['Account ID'].str.startswith('3', na=False)) & (merged_df['Account Name'].str.contains('ทุน', na=False))]
                total_capital = capital_df['TB Net Balance'].sum() if not capital_df.empty else 0.0

                if has_tax_expense:
                    applied_tax_rule = "ไม่คำนวณอัตโนมัติ (Manual)"
                    st.info("💡 **Auto-Detect:** พบรหัสบัญชี 'ภาษีเงินได้นิติบุคคล' ในงบทดลอง ระบบจะไม่คำนวณภาษีซ้ำ")
                elif total_capital <= 5000000 and total_revenue <= 30000000:
                    applied_tax_rule = "SME (ยกเว้น 300k แรก, 15%-20%)"
                    st.info(f"💡 **Auto-Detect:** ใช้งานอัตราภาษี **SME** (ทุน {total_capital:,.0f} | รายได้ {total_revenue:,.0f})")
                else:
                    applied_tax_rule = "Standard (20%)"
                    st.info(f"💡 **Auto-Detect:** ใช้งานอัตราภาษี **Standard 20%**")

            if pre_tax_profit > 0 and applied_tax_rule != "ไม่คำนวณอัตโนมัติ (Manual)":
                if applied_tax_rule == "Standard (20%)": corporate_tax = round(pre_tax_profit * 0.20, 2)
                elif applied_tax_rule == "SME (ยกเว้น 300k แรก, 15%-20%)":
                    if pre_tax_profit <= 300000: corporate_tax = 0.0
                    elif pre_tax_profit <= 3000000: corporate_tax = round((pre_tax_profit - 300000) * 0.15, 2)
                    else: corporate_tax = 405000.0 + round((pre_tax_profit - 3000000) * 0.20, 2)

            if corporate_tax > 0 and not (merged_df['Account ID'] == 'TAX-PL').any():
                tax_label = "SME" if "SME" in applied_tax_rule else "Standard"
                tax_rows = pd.DataFrame([
                    {
                        'Account ID': 'TAX-PL', 'Account Name': f'ค่าใช้จ่ายภาษีเงินได้ (Auto {tax_label})',
                        'TB Net Balance': corporate_tax, 'BS Debit': 0.0, 'BS Credit': 0.0, 'PL Debit': corporate_tax, 'PL Credit': 0.0,
                        'FS Line Item': 'ภาษีเงินได้', 'Missing in TB original': False, 'Missing in GL original': False, 'Status': 'Match'
                    },
                    {
                        'Account ID': 'TAX-BS', 'Account Name': f'ภาษีเงินได้ค้างจ่าย (Auto {tax_label})',
                        'TB Net Balance': corporate_tax, 'BS Debit': 0.0, 'BS Credit': corporate_tax, 'PL Debit': 0.0, 'PL Credit': 0.0,
                        'FS Line Item': 'สินทรัพย์หมุนเวียนอื่น', 'Missing in TB original': False, 'Missing in GL original': False, 'Status': 'Match'
                    }
                ])
                merged_df = pd.concat([merged_df, tax_rows], ignore_index=True)

            # NETTING LOGIC
            def get_fs_value(row):
                fs_line = row.get('FS Line Item', '')
                if fs_line == 'ไม่จัดประเภท (Unmapped)': return row.get('TB Net Balance', 0.0)
                rules = FS_LINE_ITEMS.get(fs_line, {})
                target_side = rules.get('side', 'bs_debit')
                bs_dr, bs_cr, pl_dr, pl_cr = row.get('BS Debit', 0.0), row.get('BS Credit', 0.0), row.get('PL Debit', 0.0), row.get('PL Credit', 0.0)
                if bs_dr == 0 and bs_cr == 0 and pl_dr == 0 and pl_cr == 0: return row.get('TB Net Balance', 0.0)

                if target_side == 'bs_debit': return bs_dr - bs_cr
                elif target_side == 'bs_credit': return bs_cr - bs_dr
                elif target_side == 'pl_debit': return pl_dr - pl_cr
                elif target_side == 'pl_credit': return pl_cr - pl_dr
                return row.get('TB Net Balance', 0.0)

            mapping_df = merged_df[['Account ID', 'Account Name', 'TB Net Balance', 'BS Debit', 'BS Credit', 'PL Debit', 'PL Credit', 'FS Line Item']].copy()
            mapping_df['FS Value'] = mapping_df.apply(get_fs_value, axis=1)
            mapping_df = mapping_df.sort_values(by=['FS Line Item', 'Account ID'])

            col_editor, col_summary = st.columns([7, 3])

            with col_editor:
                st.markdown("##### 📝 Detailed Account Mapping")
                edited_mapping = st.data_editor(
                    mapping_df,
                    column_config={
                        "FS Line Item": st.column_config.SelectboxColumn("📌 FS Line Item (Group)", options=list(FS_LINE_ITEMS.keys()) + ["ไม่จัดประเภท (Unmapped)"], required=True),
                        "Account ID": st.column_config.Column(disabled=True), "Account Name": st.column_config.Column(disabled=True),
                        "FS Value": st.column_config.NumberColumn("📊 FS Value (สุทธิ)", format="%,.2f"),
                        "BS Debit": None, "BS Credit": None, "PL Debit": None, "PL Credit": None, "TB Net Balance": None
                    },
                    use_container_width=True, height=600, key="mapping_editor"
                )

            with col_summary:
                st.markdown("##### 📈 Live FS Summary")
                summary_df = edited_mapping.groupby('FS Line Item')['FS Value'].sum().reset_index()
                def highlight_unmapped(row):
                    return ['background-color: #ffcccc'] * len(row) if row['FS Line Item'] == 'ไม่จัดประเภท (Unmapped)' and row['FS Value'] > 0 else [''] * len(row)
                st.dataframe(summary_df.style.apply(highlight_unmapped, axis=1).format({"FS Value": "{:,.2f}"}), use_container_width=True, height=600, hide_index=True)

            merged_df['FS Line Item'] = edited_mapping['FS Line Item']
            merged_df['FS Value'] = edited_mapping['FS Value']

            # ---------------------------------------------
            # HIERARCHICAL FS PREVIEW
            # ---------------------------------------------
            st.write("---")
            st.subheader("📑 Interactive Balance Sheet Preview (งบฐานะการเงิน)")
            st.write("Review the final structure. Click on line items to see the raw TB data. Highlighted rows are auto-calculated.")

            pl_summary = {k: edited_mapping.loc[edited_mapping['FS Line Item'] == k, 'FS Value'].sum() for k, v in FS_LINE_ITEMS.items() if v['side'] in ['pl_debit', 'pl_credit']}
            pl_computed = build_fs_from_mapping(pl_summary, FS_PL_STRUCTURE)
            net_profit = pl_computed.get('กำไร(ขาดทุน)สุทธิ', 0.0)

            bs_summary = {k: edited_mapping.loc[edited_mapping['FS Line Item'] == k, 'FS Value'].sum() for k, v in FS_LINE_ITEMS.items() if v['side'] in ['bs_debit', 'bs_credit']}
            bs_summary['กำไร(ขาดทุน)สะสม'] = bs_summary.get('กำไร(ขาดทุน)สะสม', 0.0) + net_profit
            
            total_paid_capital = paid_shares * paid_par
            if total_paid_capital > 0: bs_summary['ทุนเรือนหุ้น'] = total_paid_capital

            bs_computed = build_fs_from_mapping(bs_summary, FS_BS_STRUCTURE)

            for row_type, label, note, key in FS_BS_STRUCTURE:
                if row_type == 'spacer':
                    st.write("") 
                    continue
                if row_type == 'header':
                    st.markdown(f"#### 🏛️ {label}")
                    continue
                if row_type == 'subtotal':
                    val = bs_computed.get(label, 0.0)
                    if "รวมหนี้สินและส่วนของเจ้าของ" in label or label == "รวมสินทรัพย์": st.success(f"**{label}** \n### {val:,.2f} บาท")
                    else: st.info(f"**∑ {label}** : {val:,.2f}")
                    continue
                if row_type == 'item':
                    item_total = bs_computed.get(label, 0.0)
                    if key == 'ทุนเรือนหุ้น':
                        with st.expander(f"📄 **{label}** : {item_total:,.2f}"):
                            st.caption("ดึงข้อมูลอัตโนมัติจากการตั้งค่า Company Info:")
                            st.markdown(f"**ทุนจดทะเบียน (Authorized):**\n- หุ้นสามัญ {reg_shares:,.0f} หุ้น มูลค่าหุ้นละ {reg_par:,.2f} บาท (รวม {reg_shares*reg_par:,.2f} บาท)")
                            st.markdown(f"**ทุนที่ออกและเรียกชำระแล้ว (Paid-up):**\n- หุ้นสามัญ {paid_shares:,.0f} หุ้น มูลค่าหุ้นละ {paid_par:,.2f} บาท (รวม {paid_shares*paid_par:,.2f} บาท)")
                            st.info("💡 หมายเหตุ: ยอดที่นำไปคำนวณทางคณิตศาสตร์ในงบการเงินคือ **ทุนที่ออกและเรียกชำระแล้ว** เท่านั้น")
                    elif key == 'กำไร(ขาดทุน)สะสม':
                        with st.expander(f"📄 **{label}** (ยังไม่ได้จัดสรร) : {item_total:,.2f}"):
                            st.caption("Retained Earnings + Current Year Net Profit")
                            tb_re = bs_summary.get('กำไร(ขาดทุน)สะสม', 0.0) - net_profit
                            st.write(f"- กำไรสะสมต้นงวด (จาก TB): {tb_re:,.2f}")
                            st.write(f"- กำไร(ขาดทุน)สุทธิปีปัจจุบัน: {net_profit:,.2f}")
                    else:
                        with st.expander(f"📄 **{label}** : {item_total:,.2f}"):
                            mask = edited_mapping['FS Line Item'] == key
                            line_df = edited_mapping[mask]
                            if not line_df.empty:
                                display_df = line_df[['Account ID', 'Account Name', 'FS Value']].reset_index(drop=True)
                                st.dataframe(display_df, use_container_width=True, column_config={"Account ID": "รหัสบัญชี", "Account Name": "ชื่อบัญชี", "FS Value": st.column_config.NumberColumn("ยอดเงิน (Value)", format="%,.2f")})
                            else: st.warning("ยังไม่มีบัญชีที่ผูกกับรายการนี้")

        # ---------------------------------------------
        # EXPORT REPORTS
        # ---------------------------------------------
        with tab_export:
            st.write("Download your reconciled data or generate the finalized Financial Statements.")
            def compute_fs_summary(df):
                summary = {}
                for fs_line, rules in FS_LINE_ITEMS.items():
                    side = rules.get('side', 'bs_debit')
                    col_map = {'bs_debit': 'BS Debit', 'bs_credit': 'BS Credit', 'pl_debit': 'PL Debit', 'pl_credit': 'PL Credit'}
                    col = col_map.get(side, 'TB Net Balance')
                    mask = df['FS Line Item'] == fs_line
                    if col in df.columns: summary[fs_line] = df.loc[mask, col].sum()
                    else: summary[fs_line] = df.loc[mask, 'TB Net Balance'].sum()
                return summary
            
            fs_summary_current = compute_fs_summary(edited_mapping)
            fs_summary_current.pop('ไม่จัดประเภท (Unmapped)', None)
            years_data = {current_year_label: fs_summary_current}
            prior_year_arg = None
            prior_summary = {}

            if prior_tb_file:
                try:
                    prior_tb_file.seek(0)
                    prior_tb_df = parse_tb_working_paper_pdf(prior_tb_file) if prior_tb_file.name.lower().endswith('.pdf') else parse_tb(prior_tb_file)
                    if not prior_tb_df.empty:
                        prior_tb_df['FS Line Item'] = prior_tb_df.apply(auto_map, axis=1)
                        prior_summary = compute_fs_summary(prior_tb_df)
                        prior_summary.pop('ไม่จัดประเภท (Unmapped)', None)
                        years_data[prior_year_label] = prior_summary
                        prior_year_arg = prior_year_label
                        st.success(f"✅ Prior year ({prior_year_label}) loaded: {len(prior_tb_df)} accounts")
                except Exception as e: st.warning(f"Could not parse prior year TB: {e}")

            col_e1, col_e2, col_e3 = st.columns(3)
            with col_e1:
                st.markdown("##### 📋 Reconciliation Report")
                buffer_recon = io.BytesIO()
                with pd.ExcelWriter(buffer_recon, engine='xlsxwriter') as writer:
                    rename_dict = {'TB Net Balance': 'TB Net Balance (กระดาษทำการ)', 'GL Net Balance': 'GL Net Balance (บัญชีแยกประเภท)', 'TB Brought Forward Net': 'TB Brought Forward Net (งบทดลอง)', 'GL Brought Forward': 'GL Brought Forward (บัญชีแยกประเภท)', 'TB Debit': 'TB Debit (กระดาษทำการ)', 'TB Credit': 'TB Credit (กระดาษทำการ)', 'GL Debit': 'GL Debit (บัญชีแยกประเภท)', 'GL Credit': 'GL Credit (บัญชีแยกประเภท)'}
                    export_df = merged_df.drop(columns=['Missing in TB original', 'Missing in GL original']).rename(columns=rename_dict)
                    export_df.to_excel(writer, sheet_name='All Accounts', index=False)
                    mismatches_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Discrepancies', index=False)
                    missing_in_tb_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Missing in TB', index=False)
                    missing_in_gl_df.rename(columns=rename_dict).to_excel(writer, sheet_name='Missing in GL', index=False)
                    suspect_df.drop(columns=['Missing in TB original', 'Missing in GL original']).rename(columns=rename_dict).to_excel(writer, sheet_name='Top 10 Suspects', index=False)
                st.download_button(label="Download Reconciliation Data", data=buffer_recon.getvalue(), file_name="Reconciliation_Report.xlsx", mime="application/vnd.ms-excel")

            with col_e2:
                st.markdown("##### 📊 Financial Statements (New)")
                st.caption("สร้างงบการเงินจากข้อมูล FS Mapping — รองรับงบเปรียบเทียบหลายปี")
                if not company_name: st.warning("กรุณากรอกชื่อบริษัทด้านบน (Company Info) ก่อนสร้างงบ")
                else:
                    try:
                        fs_bytes = generate_fs_excel(years_data=years_data, company_name=company_name, current_year=current_year_label, prior_year=prior_year_arg, registered_capital=reg_shares * reg_par, shares=paid_shares, par_value=paid_par)
                        safe_name = re.sub(r'[^\w\u0e00-\u0e7f]', '_', company_name)[:30]
                        st.download_button(label="📊 Download Financial Statements", data=fs_bytes, file_name=f"FS_{safe_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                    except Exception as e: st.error(f"Error generating FS: {e}")

            with col_e3:
                st.markdown("##### 📄 Legacy FS Template")
                st.caption("Upload your own template. The system will search for FS Line Items and fill in the adjacent columns.")
                legacy_template_file = st.file_uploader("Upload FS Template (Excel)", type=["xls", "xlsx"], key="legacy_fs_template")
                if legacy_template_file:
                    try:
                        wb = openpyxl.load_workbook(legacy_template_file)
                        fs_names = list(FS_LINE_ITEMS.keys())
                        for sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if isinstance(cell.value, str):
                                        if "{{Company_Name}}" in cell.value: cell.value = cell.value.replace("{{Company_Name}}", company_name)
                                        if "{{Current_Year}}" in cell.value: cell.value = cell.value.replace("{{Current_Year}}", current_year_label)
                                        if prior_year_label and "{{Prior_Year}}" in cell.value: cell.value = cell.value.replace("{{Prior_Year}}", prior_year_label)
                                    if cell.column <= 5 and isinstance(cell.value, str):
                                        val_str = str(cell.value).strip()
                                        if val_str in fs_names:
                                            cy_amount = fs_summary_current.get(val_str, 0)
                                            py_amount = prior_summary.get(val_str, 0)
                                            if cy_amount != 0 or py_amount != 0:
                                                cy_cell = sheet.cell(row=cell.row, column=cell.column+1)
                                                py_cell = sheet.cell(row=cell.row, column=cell.column+2)
                                                cy_cell.value = cy_amount
                                                if prior_tb_file: py_cell.value = py_amount
                        buffer_fs = io.BytesIO()
                        wb.save(buffer_fs)
                        buffer_fs.seek(0)
                        st.download_button(label="Download Populated Template", data=buffer_fs.getvalue(), file_name=f"Generated_{legacy_template_file.name}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_legacy_fs")
                    except Exception as e: st.error(f"Error generating FS: {e}")
                else: st.info("Please upload an FS template to use this feature.")
main()