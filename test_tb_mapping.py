import pandas as pd

df = pd.read_excel("../กระดาษทำการ 2568.xls", header=None)
for index, row in df.iterrows():
    if index >= 3 and index <= 10:
        print(f"Row {index+1}: {row.tolist()}")
