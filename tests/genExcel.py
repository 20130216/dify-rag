import pandas as pd

data = {
    "日期": ["2024-07-01", "2024-07-02"],
    "项目": ["销售收入", "采购支出"],
    "金额": [10000, -5000],
    "备注": ["线上", "原材料"]
}

df = pd.DataFrame(data)
df.to_excel("tests/data/finance.xlsx", index=False)