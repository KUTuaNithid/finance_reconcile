import pandas as pd
df = pd.read_excel("../กระดาษทำการ 2568.xls", header=None)
for index, row in df.iterrows():
    acc_id = str(row[0]).strip()
    if acc_id in ['2131-11', '2135-00']:
        print(f"Row {index}: {row.tolist()}")
