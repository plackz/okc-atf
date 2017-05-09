import pandas as pd
import shutil

folder="I:\\Quality Control\\After the Fact Documentation\\ATF Sort"
 
df0 = pd.ExcelFile(folder + "\\ATF_FIFO_prev.xlsx").parse("Sheet1")
df1 = pd.ExcelFile(folder + "\\ATF_FIFO.xlsx").parse("Sheet1")
df0 = df0.set_index("ID")
df1 = df1.set_index("ID")

a0, a1 = df0.align(df1)
different = (a0 != a1).any(axis=1)
comp = a0[different].join(a1[different], lsuffix='_out', rsuffix='_in')

print(comp.head())

comp.to_excel(folder + "\\ATF_comp.xlsx")
