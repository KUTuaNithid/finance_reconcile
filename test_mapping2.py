import pandas as pd
import app

tb_df = app.parse_tb("../กระดาษทำการ 2568.xls")

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

def auto_map(row):
    acc_id = str(row['Account ID'])
    acc_name = str(row['Account Name'])
    
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

tb_df['FS Line Item'] = tb_df.apply(auto_map, axis=1)
print(tb_df[['Account ID', 'Account Name', 'FS Line Item']].to_string())
