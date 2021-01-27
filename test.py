import pandas as pd
import numpy as np
from openpyxl import load_workbook

df = pd.read_excel("Tests/test.xlsx")
x = {}

for index, row in df.iterrows():
    if row['Number'] in x.keys():
        x[row['Number']].extend(row['Letter'])
    else:
        x[row['Number']] = [row['Letter']]

# handle arrays of unequal lengths
df1 = pd.concat([pd.DataFrame(v, columns=[k]) for k, v in x.items()], axis=1)
df1 = df1.transpose()

print(df1)
