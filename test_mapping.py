import pandas as pd
import app

tb_df = app.parse_tb("../กระดาษทำการ 2568.xls")

FS_LINE_ITEMS = {
    "เงินสดและรายการเทียบเท่าเงินสด": {"prefixes": ["111"], "keywords": ["เงินสด", "เงินฝาก"]},
    "ลูกหนี้การค้า": {"prefixes": ["113"], "keywords": ["ลูกหนี้"]},
    "สินทรัพย์หมุนเวียนอื่น ": {"prefixes": ["115", "119"], "keywords": ["ภาษีถูกหัก", "จ่ายล่วงหน้า"]},
    "อุปกรณ์-สุทธิ": {"prefixes": ["141"], "keywords": ["เครื่องมือ", "เครื่องจักร", "อุปกรณ์"]},
    "เจ้าหนี้อื่น": {"prefixes": ["211", "213"], "keywords": ["เจ้าหนี้", "ค้างจ่าย", "กรมสรรพากร"]},
    "เงินกู้ยืมจากบุคคลที่เกี่ยวข้องกัน": {"prefixes": ["2138"], "keywords": ["เงินกู้ยืม"]},
    "หนี้สินหมุนเวียนอื่น ": {"prefixes": ["2131", "2132", "2135"], "keywords": ["ภาษีขาย", "ภาษีหัก"]},
    "ทุนเรือนหุ้น ": {"prefixes": ["31"], "keywords": ["ทุน"]},
    "กำไร(ขาดทุน)สะสม": {"prefixes": ["32"], "keywords": ["กำไร"]},
    "รายได้จากการให้บริการ": {"prefixes": ["41"], "keywords": ["รายได้จากการ"]},
    "รายได้อื่น": {"prefixes": ["42"], "keywords": ["รายได้อื่น", "ดอกเบี้ย"]},
    "ต้นทุนการให้บริการ": {"prefixes": ["51"], "keywords": ["ต้นทุน", "ซื้อ"]},
    "ค่าใช้จ่ายในการขายและบริหาร": {"prefixes": ["52", "53"], "keywords": ["ค่าใช้จ่าย", "เงินเดือน", "ค่าธรรมเนียม", "ค่าเสื่อม", "ค่าเช่า", "ประกัน"]},
}

def auto_map(row):
    acc_id = str(row['Account ID'])
    acc_name = str(row['Account Name'])
    
    # Prefix match (Higher priority)
    for fs_line, rules in FS_LINE_ITEMS.items():
        if any(acc_id.startswith(p) for p in rules["prefixes"]):
            return fs_line
            
    # Keyword match (Lower priority)
    for fs_line, rules in FS_LINE_ITEMS.items():
        if any(k in acc_name for k in rules["keywords"]):
            return fs_line
            
    return "ไม่จัดประเภท (Unmapped)"

tb_df['FS Line Item'] = tb_df.apply(auto_map, axis=1)
print(tb_df[['Account ID', 'Account Name', 'FS Line Item']].head(15).to_string())
