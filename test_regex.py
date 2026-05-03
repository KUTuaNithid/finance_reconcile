import re

lines = [
    '1111-00              เงินสด                                                                                                   226,153.96',
    '1111-00              เงินสด                                                     (ต\uf70aอ)',
    '2131-10              ค\uf70aาตรวจสอบบัญชีค\uf70bางจ\uf70aาย                                                                                       (8,000.00)',
    '3100-00              ทุน                                                                                                           (250,000.00)'
]

for line in lines:
    match_id = re.match(r'^(\d{4}-\d{2})\s+(.*)', line)
    if match_id:
        current_acc_id = match_id.group(1)
        rest_of_line = match_id.group(2)
        bf_match = re.search(r'\s+((?:\()?[\d,]+\.\d{2}(?:\))?)$', rest_of_line)
        if bf_match:
            bf_str = bf_match.group(1)
            is_negative = '(' in bf_str
            bf_val = float(bf_str.replace('(', '').replace(')', '').replace(',', ''))
            if is_negative:
                bf_val = -bf_val
            print(f"[{current_acc_id}] BF: {bf_val}")
        else:
            print(f"[{current_acc_id}] BF: 0.0 (No match)")
