import pandas as pd

# 1. 读取 Excel 文件
df = pd.read_excel(r'F:\python basic\sales_data.xlsx', engine='openpyxl')

# 2. 按条件筛选（筛选出 部门为'销售部' 且 销售额大于 10000 的记录）
condition = (df['部门'] == '销售部') & (df['销售额'] > 10000)
filtered_df = df[condition]

# 3. 输出保存到新的 Excel 文件
filtered_df.to_excel(r'F:\python basic\筛选结果.xlsx', index=False)

print("筛选完成！已经生成了 筛选结果.xlsx 文件")