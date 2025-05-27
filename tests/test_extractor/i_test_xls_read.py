import pandas as pd
print("Pandas version:", pd.__version__)
import xlrd
print("xlrd version:", xlrd.__version__)
df = pd.read_excel('/Users/wingzheng/Downloads/解析结果评测/dify-rag-test/2021小桃账单.xls'   , engine="xlrd")
print(df.head())