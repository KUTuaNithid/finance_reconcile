import fitz
import re
import pandas as pd

def parse_tb_pdf(pdf_path):
    doc = fitz.open(pdf_path)
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
            
            # The rest of the line has: Name, BF_Dr, BF_Cr, Movement_Dr, Movement_Cr, Ending_Dr, Ending_Cr
            # We want to extract the first two numbers.
            # Names can contain spaces, but numbers are aligned to the right.
            # Numbers can be like '226,153.96' or '0.00'
            # Let's find all sequences of digits/commas/decimals at the end of the line
            numbers = re.findall(r'((?:\()?[\d,]+\.\d{2}(?:\))?)', rest_of_line)
            if len(numbers) >= 6:
                # The first two of these 6 numbers are BF Debit and BF Credit
                # BUT what if the account name has numbers? e.g. "มาตรา 3" 
                # Our regex matches numbers with 2 decimal places, so "3" won't match. "191-1-51307-9" won't match.
                try:
                    bf_dr_str = numbers[-6].replace(',', '').replace('(', '-').replace(')', '')
                    bf_cr_str = numbers[-5].replace(',', '').replace('(', '-').replace(')', '')
                    
                    bf_dr = float(bf_dr_str)
                    bf_cr = float(bf_cr_str)
                    
                    tb_bf_data.append({
                        'Account ID': acc_id,
                        'TB Brought Forward Net': abs(bf_dr - bf_cr)
                    })
                except Exception as e:
                    print(f"Error parsing line: {line} - {e}")
                    
    return pd.DataFrame(tb_bf_data)

df = parse_tb_pdf("../งบทดลอง 2568.pdf")
print(df.head(10))
print(f"Total rows: {len(df)}")
