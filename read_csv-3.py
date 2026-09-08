import pandas as pd

# 1. 读取数据
df = pd.read_csv("数据1.csv", encoding="utf-8")

# 2. 去重（按姓名）
df_name_dedup = df.drop_duplicates(subset=['姓名'])

# 3. 填充空值
df_filled = df_name_dedup.copy()
df_filled['销售额'] = df_filled['销售额'].fillna(0)

# 4. 周四任务：按字段分组聚合
# 按"部门"分组，对"销售额"求和
grouped = df_filled.groupby('部门')['销售额'].sum().reset_index()
grouped.columns = ['部门', '总销售额']  # 给聚合结果起个清晰的列名

# 打印结果
print("===== 各部门总销售额 =====")
print(grouped)

# 5. 生成 Excel，包含多个子表
with pd.ExcelWriter('数据处理结果_含聚合.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='原始数据', index=False)
    df_name_dedup.to_excel(writer, sheet_name='按姓名去重', index=False)
    df_filled.to_excel(writer, sheet_name='填充空值后', index=False)
    grouped.to_excel(writer, sheet_name='按部门聚合', index=False)

print("✅ 处理完成！已生成 '数据处理结果_含聚合.xlsx'，包含 4 个子表。")