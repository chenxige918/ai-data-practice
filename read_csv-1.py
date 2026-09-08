import pandas as pd

# 读取数据（你刚才跑通的是 utf-8，这里保持一致）
df = pd.read_csv("数据1.csv", encoding="utf-8")

print("===== 原始数据 =====")
print(df)

# 周二任务：数据去重
# 1. 完全去重（所有列都一样才认为是重复）
df_dedup = df.drop_duplicates()
print("\n===== 完全去重后 =====")
print(df_dedup)

# 2. 按某列去重（呼应你的截图：按“姓名”去重）
# 默认 keep='first' 保留第一次出现的，keep='last' 保留最后一次出现的
df_name_dedup = df.drop_duplicates(subset=['姓名'])
print("\n===== 按姓名去重后 =====")
print(df_name_dedup)