import pandas as pd

# 读取 CSV 文件（指定 utf-8-sig 编码，避免中文乱码）
df = pd.read_csv("数据1.csv", encoding="utf-8")

# 打印整个表格
print("===== 数据预览 =====")
print(df)

# 打印基本信息（列名、非空行数、数据类型）
print("\n===== 基本信息 =====")
print(df.info())

# 打印各部门人数统计（提前试试 value_counts）
print("\n===== 各部门人数 =====")
print(df["部门"].value_counts())