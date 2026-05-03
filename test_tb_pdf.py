import fitz

with open("../งบทดลอง 2568.pdf", "rb") as f:
    doc = fitz.open(stream=f.read(), filetype="pdf")
    text = ""
    for i, page in enumerate(doc):
        text += f"--- Page {i+1} ---\n"
        text += page.get_text()
        if i >= 1:
            break
            
print(text[:2500])
