import pandas as pd

def read_data(file_path):
    """读取 Excel 文件"""
    df = pd.read_excel(file_path, engine='openpyxl')
    print(f"✅ 读取成功，共 {len(df)} 行数据")
    return df

def filter_data(df, department, min_sales):
    """
    按条件筛选
    :param df: 原始数据
    :param department: 部门名称（字符串）
    :param min_sales: 最低销售额（数字）
    """
    condition = (df['部门'] == department) & (df['销售额'] > min_sales)
    filtered_df = df[condition].copy()
    print(f"✅ 筛选完成：{department} 且 销售额>{min_sales}，共 {len(filtered_df)} 条记录")
    return filtered_df

def save_data(df, output_path):
    """保存结果到 Excel"""
    df.to_excel(output_path, index=False)
    print(f"📄 结果已保存：{output_path}")

def main():
    input_file = r'F:\python basic\sales_data.xlsx'
    
    # 场景1：销售部 > 10000
    df = read_data(input_file)
    filtered_df = filter_data(df, '销售部', 10000)
    save_data(filtered_df, r'F:\python basic\销售部_1万以上.xlsx')
    
    # 场景2：市场部 > 15000（不用改函数，只改参数）
    filtered_df2 = filter_data(df, '市场部', 15000)
    save_data(filtered_df2, r'F:\python basic\市场部_1万5以上.xlsx')
    
    # 场景3：技术部 > 8000
    filtered_df3 = filter_data(df, '技术部', 8000)
    save_data(filtered_df3, r'F:\python basic\技术部_8千以上.xlsx')
    
    print("全部完成！")

if __name__ == '__main__':
    main()