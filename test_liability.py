import fitz
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

lines = text.split('\n')
for line in lines:
    if line.startswith('21') and 'รวม' not in line:
        print(repr(line))
