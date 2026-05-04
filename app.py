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

FS_LINE_ITEMS = {
    # Balance Sheet — Assets (use BS Debit column = งบดุล เดบิต)
    "เงินสดและรายการเทียบเท่าเงินสด": {"prefixes": ["111"], "keywords": ["เงินสด", "เงินฝาก"], "side": "bs_debit"},
    "ลูกหนี้การค้า": {"prefixes": ["113"], "keywords": ["ลูกหนี้"], "side": "bs_debit"},
    "เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน": {"prefixes": ["121"], "keywords": ["เงินให้กู้ยืม"], "side": "bs_debit"},
    "สินทรัพย์หมุนเวียนอื่น ": {"prefixes": ["115", "119", "150"], "keywords": ["ภาษีถูกหัก", "จ่ายล่วงหน้า", "ดอกเบี้ยค้างรับ"], "side": "bs_debit"},
    "อุปกรณ์-สุทธิ (Gross)": {"prefixes": ["141"], "keywords": ["เครื่องมือ", "เครื่องจักร", "อุปกรณ์สำนักงาน"], "side": "bs_debit"},
    "ค่าเสื่อมราคาสะสม": {"prefixes": ["142"], "keywords": ["ค่าเสื่อมราคาสะสม"], "side": "bs_credit"},
    
    # Balance Sheet — Liabilities (use BS Credit column = งบดุล เครดิต)
    "เจ้าหนี้การค้า": {"prefixes": ["212"], "keywords": ["เจ้าหนี้การค้า"], "side": "bs_credit"},
    "เจ้าหนี้อื่น": {"prefixes": ["211", "2131"], "keywords": ["เจ้าหนี้", "ค้างจ่าย", "กรมสรรพากร", "ประกันสังคม", "สอบบัญชี", "ทำบัญชี"], "side": "bs_credit"},
    "หนี้สินหมุนเวียนอื่น ": {"prefixes": ["2132", "2137"], "keywords": ["ภาษีหัก", "ภงด"], "side": "bs_credit"},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"prefixes": ["2138"], "keywords": ["เงินกู้ยืม"], "side": "bs_credit"},
    
    # Balance Sheet — Equity (use BS Credit column = งบดุล เครดิต)
    "ทุนเรือนหุ้น ": {"prefixes": ["31"], "keywords": ["ทุน"], "side": "bs_credit"},
    "กำไร(ขาดทุน)สะสม": {"prefixes": ["32"], "keywords": ["กำไร"], "side": "bs_debit"},
    
    # P&L — Revenue (use PL Credit column = งบกำไรขาดทุน เครดิต)
    "รายได้จากการให้บริการ": {"prefixes": ["41"], "keywords": ["รายได้จากการ"], "side": "pl_credit"},
    "รายได้อื่น": {"prefixes": ["42"], "keywords": ["รายได้อื่น", "ดอกเบี้ยรับ"], "side": "pl_credit"},
    
    # P&L — Expenses (use PL Debit column = งบกำไรขาดทุน เดบิต)
    "ต้นทุนการให้บริการ": {"prefixes": ["51"], "keywords": ["ต้นทุน", "ซื้อ"], "side": "pl_debit"},
    "ค่าใช้จ่ายในการขายและบริหาร": {"prefixes": ["52", "53"], "keywords": ["ค่าใช้จ่าย", "เงินเดือน", "ค่าธรรมเนียม", "ค่าเสื่อม", "ค่าเช่า", "ประกัน", "สอบบัญชี", "บริการ"], "side": "pl_debit"},
}



# ──────────────────────────────────────────────────────────────────────────────
# FS STRUCTURE: ordered line items with note numbers and section metadata
# ──────────────────────────────────────────────────────────────────────────────
FS_BS_STRUCTURE = [
    # (type, label, note, mapped_key)
    # type: 'header' | 'item' | 'subtotal' | 'spacer'
    ('header', 'สินทรัพย์', None, None),
    ('header', 'สินทรัพย์หมุนเวียน', None, None),
    ('item',   'เงินสดและรายการเทียบเท่าเงินสด', 4, 'เงินสดและรายการเทียบเท่าเงินสด'),
    ('item',   'ลูกหนี้การค้า', 5, 'ลูกหนี้การค้า'),
    ('item',   'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน', None, 'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน'),
    ('item',   'สินทรัพย์หมุนเวียนอื่น', 6, 'สินทรัพย์หมุนเวียนอื่น '),
    ('subtotal','รวมสินทรัพย์หมุนเวียน', None, ['เงินสดและรายการเทียบเท่าเงินสด', 'ลูกหนี้การค้า', 'เงินให้กู้ยืมแก่บุคคลที่เกี่ยวข้องกัน', 'สินทรัพย์หมุนเวียนอื่น ']),
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
    ('item',   'เจ้าหนี้หมุนเวียนอื่น', 8, 'เจ้าหนี้อื่น'),
    ('item',   'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน', 9, 'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน'),
    ('item',   'หนี้สินหมุนเวียนอื่น', 10, 'หนี้สินหมุนเวียนอื่น '),
    ('subtotal','รวมหนี้สินหมุนเวียน', None, ['เจ้าหนี้การค้า', 'เจ้าหนี้อื่น', 'เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน', 'หนี้สินหมุนเวียนอื่น ']),
    ('subtotal','รวมหนี้สิน', None, ['รวมหนี้สินหมุนเวียน']),
    ('spacer',  None, None, None),
    ('header', 'ส่วนของเจ้าของ', None, None),
    ('item',   'ทุนเรือนหุ้น', None, 'ทุนเรือนหุ้น '),
    ('item',   'กำไร(ขาดทุน)สะสม', None, 'กำไร(ขาดทุน)สะสม'),
    ('subtotal','รวมส่วนของเจ้าของ', None, ['ทุนเรือนหุ้น ', 'กำไร(ขาดทุน)สะสม']),
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
    ('subtotal','กำไร(ขาดทุน)สุทธิ', None, ['รวมรายได้', '-รวมค่าใช้จ่าย']),
]


def build_fs_from_mapping(fs_mapping: dict, structure: list) -> dict:
    """Compute all line item values and subtotals from a {label: amount} mapping.
    Returns a dict {label: computed_value} for every row in structure."""
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


def _apply_cell_style(cell, bold=False, italic=False, indent=0,
                      bg=None, border_top=False, border_bottom=False,
                      number_format=None, align='left'):
    cell.font = Font(name='Cordia New', size=14, bold=bold, italic=italic)
    cell.alignment = Alignment(horizontal=align, vertical='center', indent=indent, wrap_text=False)
    if bg:
        cell.fill = PatternFill('solid', fgColor=bg)
    if number_format:
        cell.number_format = number_format
    thin = Side(style='thin')
    double = Side(style='double')
    cell.border = Border(
        top=double if border_top else None,
        bottom=double if border_bottom else None,
    )


def _write_fs_sheet(ws, structure, years_computed, year_labels,
                   company_name, statement_title, period_label):
    """Write a single FS sheet (BS or P&L) into an openpyxl worksheet."""
    NUM_FMT = '#,##0.00;[Red]-#,##0.00'
    HDR_FILL = 'DDEEFF'
    SUBTOT_FILL = 'F0F4F8'

    # ── Column widths ──
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 10
    for i, _ in enumerate(year_labels):
        col = get_column_letter(3 + i * 2)   # C, E, G …
        ws.column_dimensions[col].width = 18
        ws.column_dimensions[get_column_letter(4 + i * 2)].width = 2  # spacer

    row = 1
    # ── Header block ──
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

    # ── Year column headers ──
    ws.row_dimensions[row].height = 18
    for i, lbl in enumerate(year_labels):
        col = 3 + i * 2
        c = ws.cell(row, col, lbl)
        _apply_cell_style(c, bold=True, align='center')
    row += 1

    # ── Data rows ──
    for row_type, label, note, key in structure:
        ws.row_dimensions[row].height = 18
        if row_type == 'spacer':
            row += 1
            continue

        is_header = row_type == 'header'
        is_subtotal = row_type == 'subtotal'
        is_grand = label and label.startswith('รวม') and 'สินทรัพย์' in label and 'ไม่' not in label and 'หมุนเวียน' not in label
        indent = 0 if is_header else (1 if is_subtotal else 2)

        # Label cell
        c = ws.cell(row, 1, label)
        _apply_cell_style(c, bold=(is_header or is_subtotal), indent=indent,
                          bg=HDR_FILL if is_header else (SUBTOT_FILL if is_subtotal else None))

        # Note cell
        if note:
            nc = ws.cell(row, 2, note)
            _apply_cell_style(nc, align='center')

        # Value cells
        for i, lbl in enumerate(year_labels):
            col = 3 + i * 2
            val = years_computed[lbl].get(label) if label else None
            vc = ws.cell(row, col, val if val is not None else None)
            _apply_cell_style(vc, bold=is_subtotal,
                              bg=SUBTOT_FILL if is_subtotal else None,
                              number_format=NUM_FMT, align='right',
                              border_top=is_subtotal,
                              border_bottom=(is_subtotal and 'รวม' in (label or '') and 'ทั้งหมด' not in (label or '')))
        row += 1

    # Footer
    ws.row_dimensions[row].height = 14
    ws.cell(row, 1, 'หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้')
    _apply_cell_style(ws.cell(row, 1), italic=True)


def generate_fs_excel(years_data: dict, company_name: str,
                      current_year: str, prior_year: str = None,
                      registered_capital: float = 0,
                      shares: int = 0, par_value: float = 0) -> bytes:
    """
    Generate a clean standalone Financial Statement Excel workbook.

    years_data: {year_label: {FS_LINE_ITEM: amount}}
    Returns bytes of the .xlsx file.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    year_labels = [current_year]
    if prior_year and prior_year in years_data:
        year_labels.append(prior_year)

    # ── Compute BS and P&L for each year ──
    bs_computed = {}   # {year_label: {row_label: value}}
    pl_computed = {}   # {year_label: {row_label: value}}
    for yl in year_labels:
        mapping = years_data.get(yl, {})
        pl_vals = build_fs_from_mapping(mapping, FS_PL_STRUCTURE)
        # Add net profit into retained earnings for BS
        net_profit = pl_vals.get('กำไร(ขาดทุน)สุทธิ', 0.0) or 0.0
        mapping_bs = dict(mapping)
        mapping_bs['กำไร(ขาดทุน)สะสม'] = mapping.get('กำไร(ขาดทุน)สะสม', 0.0) + net_profit
        bs_computed[yl] = build_fs_from_mapping(mapping_bs, FS_BS_STRUCTURE)
        pl_computed[yl] = pl_vals

    # ── Sheet 1: งบฐานะการเงิน ──
    ws_bs = wb.create_sheet('งบฐานะการเงิน')
    _write_fs_sheet(ws_bs, FS_BS_STRUCTURE, bs_computed, year_labels,
                    company_name,
                    'งบฐานะการเงิน',
                    f'ณ วันที่ 31 ธันวาคม {current_year}')

    # ── Sheet 2: งบกำไรขาดทุน ──
    ws_pl = wb.create_sheet('งบกำไรขาดทุน')
    _write_fs_sheet(ws_pl, FS_PL_STRUCTURE, pl_computed, year_labels,
                    company_name,
                    'งบกำไรขาดทุน',
                    f'สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม {current_year}')

    # ── Sheet 3: งบส่วนของเจ้าของ ──
    _write_equity_sheet(wb, company_name, current_year, prior_year,
                        bs_computed, pl_computed, year_labels,
                        registered_capital, shares, par_value)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _write_equity_sheet(wb, company_name, current_year, prior_year,
                        bs_computed, pl_computed, year_labels,
                        registered_capital, shares, par_value):
    """Write งบการเปลี่ยนแปลงส่วนของเจ้าของ sheet."""
    ws = wb.create_sheet('งบส่วนของเจ้าของ')
    ws.column_dimensions['A'].width = 46
    ws.column_dimensions['B'].width = 2
    ws.column_dimensions['C'].width = 2
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 2
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 2
    ws.column_dimensions['H'].width = 18
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
    r += 1  # spacer

    # Column headers
    wc(r, 1, None)
    wc(r, 4, 'ทุนที่ออกและ', bold=True, align='center')
    wc(r, 6, 'กำไร(ขาดทุน)', bold=True, align='center')
    wc(r, 8, 'รวมส่วนของ', bold=True, align='center'); r += 1
    wc(r, 4, 'เรียกชำระแล้ว', bold=True, align='center')
    wc(r, 6, 'สะสม', bold=True, align='center')
    wc(r, 8, 'เจ้าของ', bold=True, align='center'); r += 1

    # Get values
    def get_pl(yl, key):
        return pl_computed.get(yl, {}).get(key, 0.0) or 0.0
    def get_bs(yl, key):
        return bs_computed.get(yl, {}).get(key, 0.0) or 0.0

    if prior_year and prior_year in bs_computed:
        # Prior year opening (assume zero start if only 1 year)
        prior_capital = get_bs(prior_year, 'ทุนเรือนหุ้น') or (shares * par_value)
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
        cur_capital_open = shares * par_value if shares and par_value else 0
        cur_retained_open = 0

    cur_net = get_pl(current_year, 'กำไร(ขาดทุน)สุทธิ')
    cur_capital = get_bs(current_year, 'ทุนเรือนหุ้น') or (shares * par_value)

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

st.set_page_config(page_title="Finance Reconcile", layout="wide", page_icon="📊")


def parse_tb(excel_file):
    # Read the excel file, skipping the first 4 rows to get to the data
    df = pd.read_excel(excel_file, header=None)
    
    # The กระดาษทำการ has 8 columns:
    # col0=เลขที่บัญชี, col1=ชื่อบัญชี
    # col2=งบทดลอง เดบิต, col3=งบทดลอง เครดิต
    # col4=งบกำไรขาดทุน เดบิต, col5=งบกำไรขาดทุน เครดิต
    # col6=งบดุล เดบิต, col7=งบดุล เครดิต
    
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
            
            # Trial Balance columns (col 2-3)
            debit = float(row[2]) if not pd.isna(row[2]) else 0.0
            credit = float(row[3]) if not pd.isna(row[3]) else 0.0
            net_balance = abs(debit - credit)
            
            # P&L columns (col 4-5)
            pl_debit = float(row[4]) if len(row) > 4 and not pd.isna(row[4]) else 0.0
            pl_credit = float(row[5]) if len(row) > 5 and not pd.isna(row[5]) else 0.0
            
            # BS columns (col 6-7)
            bs_debit = float(row[6]) if len(row) > 6 and not pd.isna(row[6]) else 0.0
            bs_credit = float(row[7]) if len(row) > 7 and not pd.isna(row[7]) else 0.0
            
            tb_data.append({
                'Account ID': acc_id,
                'Account Name': acc_name,
                'TB Debit': debit,
                'TB Credit': credit,
                'TB Net Balance': net_balance,
                'PL Debit': pl_debit,
                'PL Credit': pl_credit,
                'BS Debit': bs_debit,
                'BS Credit': bs_credit,
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
      1. งบทดลอง   (TB)  — The full trial balance
      2. งบกำไรขาดทุน (P&L) — P&L accounts only
      3. งบดุล       (BS)  — Balance sheet accounts only
    
    Each account line has numbers in the TB columns AND in either the
    BS columns (for accounts 1xxx-3xxx) or the P&L columns (for 4xxx-5xxx).
    """
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    
    lines = text.split('\n')
    tb_data = []
    
    for raw_line in lines:
        line = raw_line.rstrip('\r\n')
        
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
        
        # Extract account name: text between the account ID and the first number
        name_match = re.match(r'^(.+?)(?=\s{3,}[\d,]+\.\d{2})', rest)
        if name_match:
            acc_name = _clean_thai_pdf_text(name_match.group(1).strip())
            if re.match(r'^[\d,]+\.\d{2}$', acc_name):
                acc_name = "Unknown"
        else:
            acc_name = "Unknown"
        
        # Parse numbers based on account type
        # TB Debit/Credit are always the first number(s)
        # The LAST number is the BS or PL amount
        
        # Determine if this is a debit-normal or credit-normal account
        is_bs_account = acc_id[0] in ('1', '2', '3')
        is_debit_normal = acc_id[0] in ('1', '5')  # Assets and Expenses
        
        # TB columns
        tb_debit = 0.0
        tb_credit = 0.0
        pl_debit = 0.0
        pl_credit = 0.0
        bs_debit = 0.0
        bs_credit = 0.0
        
        if len(all_nums) >= 1:
            val = float(all_nums[0].replace(',', ''))
            if is_debit_normal:
                tb_debit = val
            else:
                tb_credit = val
        
        # If there are 2+ numbers, the last number is the BS or PL value
        if len(all_nums) >= 2:
            last_val = float(all_nums[-1].replace(',', ''))
            if is_bs_account:
                if is_debit_normal:  # Assets (1xxx)
                    bs_debit = last_val
                else:  # Liabilities (2xxx), Equity (3xxx)
                    bs_credit = last_val
            else:
                # P&L accounts (4xxx = Revenue = Credit, 5xxx = Expense = Debit)
                if acc_id[0] == '4':
                    pl_credit = last_val
                else:
                    pl_debit = last_val
        elif len(all_nums) == 1:
            # Only one number - it's both the TB and the BS/PL value
            val = float(all_nums[0].replace(',', ''))
            if is_bs_account:
                if is_debit_normal:
                    bs_debit = val
                else:
                    bs_credit = val
            else:
                if acc_id[0] == '4':
                    pl_credit = val
                else:
                    pl_debit = val
        
        net_balance = abs(tb_debit - tb_credit)
        
        tb_data.append({
            'Account ID': acc_id,
            'Account Name': acc_name,
            'TB Debit': tb_debit,
            'TB Credit': tb_credit,
            'TB Net Balance': net_balance,
            'PL Debit': pl_debit,
            'PL Credit': pl_credit,
            'BS Debit': bs_debit,
            'BS Credit': bs_credit,
        })
    
    return pd.DataFrame(tb_data)

def parse_gl(pdf_file):
    # Read PDF text
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
        
    lines = text.split('\n')
    
    # Extract Company Name and Year
    extracted_company = ""
    extracted_year = ""
    for i in range(min(10, len(lines))):
        line = lines[i].strip()
        if "บริษัท" in line and not extracted_company:
            # Capture everything up to "หนา" or a bunch of spaces
            match_comp = re.split(r'\s{3,}|\t|หน้า|หนา', line)
            extracted_company = match_comp[0].strip()
        if "วันที่จาก" in line and not extracted_year:
            # e.g., วันที่จาก       1 ม.ค. 2568      ถึง   31 ธ.ค. 2568
            match_yr = re.search(r'(25\d{2})', line)
            if match_yr:
                extracted_year = "พ.ศ. " + match_yr.group(1)
                
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
                    
    return pd.DataFrame(gl_data), extracted_company, extracted_year



def main():
    st.title("GL & TB Reconciliation Tool")
    st.markdown("Upload your General Ledger (PDF), Trial Balance (Excel), and optionally Trial Balance (PDF) to verify Brought Forward balances.")
    
    # Use extracted GL values as widget defaults (set BEFORE widgets render)
    default_company = st.session_state.get('_gl_extracted_company', '')
    default_year = st.session_state.get('_gl_extracted_year', 'พ.ศ. 2568')
    default_prior_year = st.session_state.get('_gl_extracted_prior_year', 'พ.ศ. 2567')
    
    # Company info for FS generation
    with st.expander("ข้อมูลบริษัท / Company Info (สำหรับสร้างงบการเงิน)", expanded=False):
        col_ci1, col_ci2, col_ci3, col_ci4 = st.columns([3, 1, 1, 1])
        with col_ci1:
            company_name = st.text_input("ชื่อบริษัท", value=default_company, placeholder="บริษัท xxxxxxx จำกัด", key='company_name')
        with col_ci2:
            current_year_label = st.text_input("ปีปัจจุบัน", value=default_year, key='current_year_label')
        with col_ci3:
            fs_shares = st.number_input("จำนวนหุ้น", min_value=0, value=0, step=1000, key='fs_shares')
        with col_ci4:
            fs_par = st.number_input("มูลค่าหุ้นละ (บาท)", min_value=0.0, value=0.0, step=1.0, key='fs_par')
    
    col1, col2, col3 = st.columns(3)
    with col1:
        tb_file = st.file_uploader("Upload Trial Balance / กระดาษทำการ (Excel or PDF)", type=["xls", "xlsx", "pdf"])
    with col2:
        gl_file = st.file_uploader("Upload General Ledger / บัญชีแยกประเภท (PDF)", type=["pdf"])
    with col3:
        tb_pdf_file = st.file_uploader("Upload Trial Balance for BF Check / งบทดลอง (PDF)", type=["pdf"])
    
    # Prior year TB upload for FS generation
    with st.expander("📊 อัปโหลดกระดาษทำการปีก่อนหน้า (สำหรับสร้างงบเปรียบเทียบ)", expanded=False):
        col_py1, col_py2 = st.columns([1, 2])
        with col_py1:
            prior_year_label = st.text_input("ปีก่อนหน้า", value=default_prior_year, key='prior_year_label')
        with col_py2:
            prior_tb_file = st.file_uploader("Upload Prior Year กระดาษทำการ (Excel or PDF)", type=["xls", "xlsx", "pdf"], key='prior_tb_upload')
        
    run_clicked = st.button("Run Reconciliation", type="primary")
    should_process = False
    if run_clicked:
        if not tb_file or not gl_file:
            st.error("Please upload both files first.")
            return
        should_process = True

    if should_process:
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
                gl_df, extracted_company, extracted_year = parse_gl(gl_file)
                # Store extracted values for widget defaults on next render
                if extracted_company:
                    st.session_state['_gl_extracted_company'] = extracted_company
                    # Override local variable so FS generation works this run
                    if not company_name:
                        company_name = extracted_company
                if extracted_year:
                    st.session_state['_gl_extracted_year'] = extracted_year
                    if current_year_label == 'พ.ศ. 2568' or not current_year_label:
                        current_year_label = extracted_year
                    # Calculate prior year automatically
                    try:
                        yr_int = int(re.search(r'\d{4}', extracted_year).group())
                        prior_yr = f"พ.ศ. {yr_int - 1}"
                        st.session_state['_gl_extracted_prior_year'] = prior_yr
                        if prior_year_label == 'พ.ศ. 2567' or not prior_year_label:
                            prior_year_label = prior_yr
                    except:
                        pass
                st.info(f"📋 ข้อมูลจาก GL: **{extracted_company}** | ปี: **{extracted_year}**")
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
            # New columns from กระดาษทำการ
            for col in ['PL Debit', 'PL Credit', 'BS Debit', 'BS Credit']:
                if col not in merged_df.columns:
                    merged_df[col] = 0.0
                merged_df[col] = merged_df[col].fillna(0)
            
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

        # Display Results
        st.success("Reconciliation Complete!")

        st.subheader("Summary")
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Total Accounts", len(merged_df))
        col_s2.metric("Matched Accounts", len(matches_df))
        col_s3.metric("Mismatches", len(mismatches_df))
        col_s4.metric("Missing Accounts", len(missing_in_gl_df) + len(missing_in_tb_df))

        # Apply initial mapping


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
            st.info("💡 **FS Value** คือตัวเลขที่จะถูกนำไปใช้สร้างงบการเงิน (จากงบดุล/งบกำไรขาดทุน ตามประเภทบัญชี) — สามารถแก้ไขได้ทั้ง FS Line Item และ FS Value")

            # Compute the FS Value column: the correct amount that will be used in FS generation
            def get_fs_value(row):
                fs_line = row.get('FS Line Item', '')
                rules = FS_LINE_ITEMS.get(fs_line, {})
                side = rules.get('side', 'bs_debit')
                col_map = {
                    'bs_debit': 'BS Debit',
                    'bs_credit': 'BS Credit',
                    'pl_debit': 'PL Debit',
                    'pl_credit': 'PL Credit',
                }
                col = col_map.get(side, 'TB Net Balance')
                return row.get(col, row.get('TB Net Balance', 0.0))
            
            mapping_df = merged_df[['Account ID', 'Account Name', 'TB Net Balance', 'BS Debit', 'BS Credit', 'PL Debit', 'PL Credit', 'FS Line Item']].copy()
            mapping_df['FS Value'] = merged_df.apply(get_fs_value, axis=1)
            
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
                    "TB Net Balance": st.column_config.NumberColumn("TB Net Balance", format="%,.2f"),
                    "BS Debit": st.column_config.NumberColumn("งบดุล เดบิต", format="%,.2f"),
                    "BS Credit": st.column_config.NumberColumn("งบดุล เครดิต", format="%,.2f"),
                    "PL Debit": st.column_config.NumberColumn("งบกำไรขาดทุน เดบิต", format="%,.2f"),
                    "PL Credit": st.column_config.NumberColumn("งบกำไรขาดทุน เครดิต", format="%,.2f"),
                    "FS Value": st.column_config.NumberColumn("📊 FS Value (ใช้สร้างงบ)", format="%,.2f"),
                },
                disabled=["Account ID", "Account Name", "TB Net Balance", "FS Value"],
                use_container_width=True,
                key="mapping_editor"
            )

            # Update merged_df with the edited mapping
            merged_df['FS Line Item'] = edited_mapping['FS Line Item']
            # Also update the editable BS/PL columns back
            for col in ['BS Debit', 'BS Credit', 'PL Debit', 'PL Credit']:
                merged_df[col] = edited_mapping[col]

            # Show live preview of the generated FS using correct columns
            st.subheader("Financial Statement Preview")
            st.write("Click on any line item to see the accounts that make up its total.")
            
            for fs_line, rules in FS_LINE_ITEMS.items():
                side = rules.get('side', 'bs_debit')
                col_map = {'bs_debit': 'BS Debit', 'bs_credit': 'BS Credit', 'pl_debit': 'PL Debit', 'pl_credit': 'PL Credit'}
                col = col_map.get(side, 'TB Net Balance')
                mask = merged_df['FS Line Item'] == fs_line
                line_df = merged_df[mask]
                if line_df.empty:
                    continue
                total = line_df[col].sum() if col in line_df.columns else line_df['TB Net Balance'].sum()
                side_label = {"bs_debit": "งบดุล เดบิต", "bs_credit": "งบดุล เครดิต", "pl_debit": "งบกำไรขาดทุน เดบิต", "pl_credit": "งบกำไรขาดทุน เครดิต"}.get(side, "")
                with st.expander(f"**{fs_line}** — Total: **{total:,.2f}** ({side_label})"):
                    display_cols = ['Account ID', 'Account Name', col]
                    st.dataframe(line_df[display_cols].reset_index(drop=True), use_container_width=True, column_config={col: num_config})

            # Show unmapped items if any
            unmapped = merged_df[merged_df['FS Line Item'] == "ไม่จัดประเภท (Unmapped)"]
            if not unmapped.empty:
                unmapped_total = unmapped['TB Net Balance'].sum()
                with st.expander(f"⚠️ **ไม่จัดประเภท (Unmapped)** — Total: **{unmapped_total:,.2f}**"):
                    st.dataframe(unmapped[['Account ID', 'Account Name', 'TB Net Balance']].reset_index(drop=True), use_container_width=True, column_config={'TB Net Balance': num_config})


        with tab_export:
            st.write("Download your reconciled data or generate the finalized Financial Statements.")

            # Current year FS data from the mapping — use correct column per FS Line Item
            def compute_fs_summary(df):
                """Compute FS summary using the correct column for each FS line item."""
                summary = {}
                for fs_line, rules in FS_LINE_ITEMS.items():
                    side = rules.get('side', 'bs_debit')
                    # Map side to the correct DataFrame column
                    col_map = {
                        'bs_debit': 'BS Debit',
                        'bs_credit': 'BS Credit',
                        'pl_debit': 'PL Debit',
                        'pl_credit': 'PL Credit',
                    }
                    col = col_map.get(side, 'TB Net Balance')
                    mask = df['FS Line Item'] == fs_line
                    if col in df.columns:
                        summary[fs_line] = df.loc[mask, col].sum()
                    else:
                        # Fallback to TB Net Balance if column not available
                        summary[fs_line] = df.loc[mask, 'TB Net Balance'].sum()
                return summary
            
            fs_summary_current = compute_fs_summary(edited_mapping)
            # Remove unmapped
            fs_summary_current.pop('ไม่จัดประเภท (Unmapped)', None)

            # Prepare years_data dict
            years_data = {current_year_label: fs_summary_current}
            prior_year_arg = None
            prior_summary = {}

            # Parse prior year TB if uploaded
            if prior_tb_file:
                try:
                    prior_tb_file.seek(0)
                    if prior_tb_file.name.lower().endswith('.pdf'):
                        prior_tb_df = parse_tb_working_paper_pdf(prior_tb_file)
                    else:
                        prior_tb_df = parse_tb(prior_tb_file)
                    if not prior_tb_df.empty:
                        prior_tb_df['FS Line Item'] = prior_tb_df.apply(auto_map, axis=1)
                        prior_summary = compute_fs_summary(prior_tb_df)
                        prior_summary.pop('ไม่จัดประเภท (Unmapped)', None)
                        years_data[prior_year_label] = prior_summary
                        prior_year_arg = prior_year_label
                        st.success(f"✅ Prior year ({prior_year_label}) loaded: {len(prior_tb_df)} accounts")
                except Exception as e:
                    st.warning(f"Could not parse prior year TB: {e}")

            col_e1, col_e2, col_e3 = st.columns(3)
            with col_e1:
                st.markdown("##### 📋 Reconciliation Report")
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
                st.markdown("##### 📊 Financial Statements (New)")
                st.caption("สร้างงบการเงินจากข้อมูล FS Mapping — รองรับงบเปรียบเทียบหลายปี")

                if not company_name:
                    st.warning("กรุณากรอกชื่อบริษัทด้านบน (Company Info) ก่อนสร้างงบ")
                else:
                    try:
                        fs_bytes = generate_fs_excel(
                            years_data=years_data,
                            company_name=company_name,
                            current_year=current_year_label,
                            prior_year=prior_year_arg,
                            registered_capital=fs_shares * fs_par,
                            shares=fs_shares,
                            par_value=fs_par,
                        )
                        safe_name = re.sub(r'[^\w\u0e00-\u0e7f]', '_', company_name)[:30]
                        st.download_button(
                            label="📊 Download Financial Statements",
                            data=fs_bytes,
                            file_name=f"FS_{safe_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                    except Exception as e:
                        st.error(f"Error generating FS: {e}")

            with col_e3:
                st.markdown("##### 📄 Legacy FS Template")
                st.caption("Upload your own template. The system will search for FS Line Items and fill in the adjacent columns.")
                legacy_template_file = st.file_uploader("Upload FS Template (Excel)", type=["xls", "xlsx"], key="legacy_fs_template")

                if legacy_template_file:
                    try:
                        wb = openpyxl.load_workbook(legacy_template_file)
                        fs_names = list(FS_LINE_ITEMS.keys())
                        # Iterate all sheets to replace placeholders and find row labels
                        for sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            for row in sheet.iter_rows():
                                for cell in row:
                                    # Replace placeholder tags
                                    if isinstance(cell.value, str):
                                        if "{{Company_Name}}" in cell.value:
                                            cell.value = cell.value.replace("{{Company_Name}}", company_name)
                                        if "{{Current_Year}}" in cell.value:
                                            cell.value = cell.value.replace("{{Current_Year}}", current_year_label)
                                        if prior_year_label and "{{Prior_Year}}" in cell.value:
                                            cell.value = cell.value.replace("{{Prior_Year}}", prior_year_label)

                                    # Search for FS Line Items (scan first 5 columns)
                                    if cell.column <= 5 and isinstance(cell.value, str):
                                        val_str = str(cell.value).strip()
                                        if val_str in fs_names:
                                            # Found a match! Current year goes to column+1, Prior year to column+2
                                            cy_amount = fs_summary_current.get(val_str, 0)
                                            py_amount = prior_summary.get(val_str, 0)
                                            
                                            # Only write if there is an amount or if it's safe to zero it out
                                            if cy_amount != 0 or py_amount != 0:
                                                cy_cell = sheet.cell(row=cell.row, column=cell.column+1)
                                                py_cell = sheet.cell(row=cell.row, column=cell.column+2)
                                                
                                                # Set values
                                                cy_cell.value = cy_amount
                                                if prior_tb_file:
                                                    py_cell.value = py_amount

                        buffer_fs = io.BytesIO()
                        wb.save(buffer_fs)
                        buffer_fs.seek(0)
                        st.download_button(
                            label="Download Populated Template",
                            data=buffer_fs.getvalue(),
                            file_name=f"Generated_{legacy_template_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_legacy_fs"
                        )
                    except Exception as e:
                        st.error(f"Error generating FS: {e}")
                else:
                    st.info("Please upload an FS template to use this feature.")
main()
