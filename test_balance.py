import fitz
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for page in doc[:3]:
        text += page.get_text()

lines = text.split('\n')
for i, line in enumerate(lines[:100]):
    if '1111-00' in line:
        print(f"Line {i}: {repr(line)}")
