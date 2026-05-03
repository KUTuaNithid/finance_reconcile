import app
with open("../บัญชีแยกประเภท 2568.pdf", "rb") as f:
    gl_df = app.parse_gl(f)
print(gl_df[gl_df['Account ID'] == '2131-11'])
print(gl_df[gl_df['Account ID'] == '2135-00'])
