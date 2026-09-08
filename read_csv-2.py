import pandas as pd

# 1. 读取数据
df = pd.read_csv("数据1.csv", encoding="utf-8")

# 2. 去重操作
df_name_dedup = df.drop_duplicates(subset=['姓名'])

# 3. 填充空值
df_filled = df_name_dedup.copy()
df_filled['销售额'] = df_filled['销售额'].fillna(0)

# 调试打印
print("===== 填充空值后的数据 =====")
print(df_filled)

# 4. 生成新的 Excel 文件
with pd.ExcelWriter('数据处理结果.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='原始数据', index=False)
    df_name_dedup.to_excel(writer, sheet_name='按姓名去重', index=False)
    df_filled.to_excel(writer, sheet_name='填充空值后', index=False)

print("✅ 处理完成！已生成 '数据处理结果.xlsx'")