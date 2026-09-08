import pandas as pd
from openpyxl.styles import PatternFill

# 1. 读取 Excel (跳过第一行大标题)
file_path = r'F:\python basic\FS.xlsx'
df = pd.read_excel(file_path, engine='openpyxl', skiprows=1)

# 2. 定义红色背景
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# 3. 创建输出文件
output_path = r'F:\python basic\FS_标红结果.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='PK榜', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['PK榜']
    
    # 4. 遍历每一行，标红 PK 胜方
    for row in range(2, 10):
        cell_left = worksheet.cell(row=row, column=3)  # 左边分数(C列)
        cell_right = worksheet.cell(row=row, column=6) # 右边分数(F列)
        
        is_left_win = False
        
        # 确保左边是数字
        if isinstance(cell_left.value, (int, float)):
            # 如果右边是空值（第一行情况），默认左边赢
            if cell_right.value is None or cell_right.value == "":
                is_left_win = True
            # 如果右边也是数字，比较大小
            elif isinstance(cell_right.value, (int, float)):
                if cell_left.value > cell_right.value:
                    is_left_win = True
                    
        # 执行标红操作
        if is_left_win:
            for col in range(2, 4): # 标红左边整行(A, B, C)
                worksheet.cell(row=row, column=col).fill = red_fill
        else:
            # 右边赢的情况
            if isinstance(cell_right.value, (int, float)) and isinstance(cell_left.value, (int, float)):
                if cell_right.value > cell_left.value:
                    for col in range(5, 7): # 标红右边整行(D, E, F)
                        worksheet.cell(row=row, column=col).fill = red_fill

print("PK胜方标红完成！生成了 FS_标红结果.xlsx")