import pandas as pd

# 读取数据
df = pd.read_csv("数据1.csv", encoding="utf-8")

# 去重操作
df_dedup = df.drop_duplicates()
df_name_dedup = df.drop_duplicates(subset=['姓名'])

# 写入同一个 Excel 文件，分不同子表
with pd.ExcelWriter('去重结果.xlsx', engine='openpyxl') as writer:
    df_dedup.to_excel(writer, sheet_name='完全去重', index=False)
    df_name_dedup.to_excel(writer, sheet_name='按姓名去重', index=False)

print("✅ Excel 文件生成成功！包含 2 个子表：'完全去重' 和 '按姓名去重'")