import pandas as pd
import fitz
import re
import app

tb_df = app.parse_tb("../กระดาษทำการ 2568.xls")
print("TB Extracted accounts:", len(tb_df))

with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    gl_df = app.parse_gl(f)
print("GL Extracted accounts:", len(gl_df))

if len(gl_df) > 0:
    print(gl_df.head(5))
