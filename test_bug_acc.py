import fitz
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '1410-07' in line or '2138-00' in line or '3100-00' in line:
        print(f"Line {i}: {repr(line)}")
        for j in range(i, i+10):
            if 'รวม' in lines[j]:
                print(f"  Total line {j}: {repr(lines[j])}")
                break
