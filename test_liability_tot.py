import fitz
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()

lines = text.split('\n')
for i, line in enumerate(lines):
    if '2131-10' in line:
        print(repr(line))
        for j in range(i, i+15):
            if 'รวม' in lines[j]:
                print(repr(lines[j]))
                break
