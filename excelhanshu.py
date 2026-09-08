import pandas as pd

def read_data(file_path):
    """读取 Excel 文件"""
    df = pd.read_excel(file_path, engine='openpyxl')
    print(f"✅ 读取成功，共 {len(df)} 行数据")
    return df

def filter_data(df):
    """按条件筛选：部门为销售部 且 销售额大于10000"""
    condition = (df['部门'] == '销售部') & (df['销售额'] > 10000)
    filtered_df = df[condition].copy()
    print(f"✅ 筛选完成，符合条件的有 {len(filtered_df)} 条记录")
    return filtered_df

def save_data(df, output_path):
    """保存结果到 Excel"""
    df.to_excel(output_path, index=False)
    print(f"📄 结果已保存：{output_path}")

def main():
    input_file = r'F:\python basic\sales_data.xlsx'
    output_file = r'F:\python basic\筛选结果1.xlsx'
    
    # 三步流程，清晰明了
    df = read_data(input_file)
    filtered_df = filter_data(df)
    save_data(filtered_df, output_file)
    
    print("全部完成！")

if __name__ == '__main__':
    main()