import fitz
import re

with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '1113-02' in line:
        print(f"Line {i}: {repr(line)}")
        for j in range(i, i+15):
            if 'รวม' in lines[j]:
                print(f"  Total line {j}: {repr(lines[j])}")
                break

import app
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    gl_df = app.parse_gl(f)
print("\nDataFrame row:")
print(gl_df[gl_df['Account ID'] == '1113-02'])

