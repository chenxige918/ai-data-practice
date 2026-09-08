import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def read_excel_and_calculate(input_file):
    """读取Excel并计算借贷差额"""
    # 表头在第1行，所以 header=0
    df = pd.read_excel(input_file, header=0)
    df.columns = df.columns.str.strip()
    
    print("当前Excel读取到的列名：")
    print(df.columns.tolist())
    
    # 智能匹配借贷列（支持多组借方/贷方）
    debit_cols = [c for c in df.columns if c == '借方']
    credit_cols = [c for c in df.columns if c == '贷方']
    
    if not debit_cols or not credit_cols:
        print("❌ 没找到'借方'或'贷方'列！")
        return None
    
    # 科目列明确为前两列
    col_a = '科目编号'
    col_b = '科目名称'
    
    print(f"使用的科目列：{col_a}, {col_b}")
    print(f"找到的借方列：{debit_cols}，贷方列：{credit_cols}")
    
    # 计算
    df['借方合计'] = pd.to_numeric(df[debit_cols].sum(axis=1), errors='coerce').fillna(0)
    df['贷方合计'] = pd.to_numeric(df[credit_cols].sum(axis=1), errors='coerce').fillna(0)
    df['差额'] = df['借方合计'] - df['贷方合计']
    
    print(f"✅ 读取成功，共 {len(df)} 行数据")
    return df

def find_unbalanced(df):
    """筛选不平衡的科目"""
    unbalanced = df[df['差额'].abs() > 0.01].copy()
    
    def reason(x):
        if x > 0:
            return f'借方大于贷方，相差 {abs(x):.2f}'
        else:
            return f'贷方大于借方，相差 {abs(x):.2f}'
    
    unbalanced['相差原因'] = unbalanced['差额'].apply(reason)
    print(f"✅ 找到 {len(unbalanced)} 个不平衡科目")
    return unbalanced

def save_files(unbalanced, input_file, script_dir):
    """保存明细和标注文件（绝对路径）"""
    # 保存明细
    detail_path = os.path.join(script_dir, '不平衡明细.xlsx')
    unbalanced[['科目编号', '科目名称', '借方合计', '贷方合计', '差额', '相差原因']].to_excel(detail_path, index=False)
    print(f"📄 明细文件：{detail_path}")
    
    # 标红原表
    wb = load_workbook(input_file)
    ws = wb.active
    red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    ids = set(unbalanced['科目编号'].astype(str).str.strip())
    
    for row in range(2, ws.max_row + 1):
        val = str(ws.cell(row, 1).value).strip()
        if val in ids:
            ws.cell(row, 1).fill = red
            ws.cell(row, 2).fill = red
    
    marked_path = os.path.join(script_dir, '标注不平衡.xlsx')
    wb.save(marked_path)
    print(f"📄 标注文件：{marked_path}")

def main():
    # 获取 KY.py 所在的文件夹绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(script_dir, '科目余额202606.xlsx')
    
    print(f"📂 脚本目录：{script_dir}")
    print(f"📂 输入文件：{input_file}")
    
    if not os.path.exists(input_file):
        print("❌ 找不到 科目余额202606.xlsx！请确认文件和 KY.py 在同一个文件夹里。")
        return
    
    df = read_excel_and_calculate(input_file)
    if df is None:
        return
    
    unbalanced = find_unbalanced(df)
    if len(unbalanced) == 0:
        print("🎉 所有科目借贷平衡，没有生成文件。")
        return
    
    save_files(unbalanced, input_file, script_dir)
    print("✅ 全部完成！请复制上面的路径到文件资源管理器中打开。")

if __name__ == '__main__':
    main()