import pandas as pd

# 全局文件路径
INPUT_FILE = "数据1.csv"
OUTPUT_FILE = "数据处理最终结果.xlsx"

def load_data(file_path):
    """任务1：读取CSV文件"""
    print("📥 [1/4] 正在读取数据...")
    df = pd.read_csv(file_path, encoding="utf-8")
    print(f"读取完成，共 {len(df)} 行数据。")
    return df

def remove_duplicates(df):
    """任务2：按姓名去重"""
    print("🧹 [2/4] 正在按姓名去重...")
    df_dedup = df.drop_duplicates(subset=['姓名'])
    print(f"去重完成，剩余 {len(df_dedup)} 行数据。")
    return df_dedup

def fill_missing_values(df):
    """任务3：填充空值"""
    print("🔧 [3/4] 正在填充空值...")
    df_filled = df.copy()
    df_filled['销售额'] = df_filled['销售额'].fillna(0)
    print("空值填充完成（销售额空值已填为0）。")
    return df_filled

def group_aggregate(df):
    """任务4：按字段分组聚合"""
    print("📊 [4/4] 正在按部门分组聚合...")
    grouped = df.groupby('部门')['销售额'].sum().reset_index()
    grouped.columns = ['部门', '总销售额']
    print("聚合完成：")
    print(grouped)
    return grouped

def export_to_excel(data_frames, sheet_names):
    """导出到同一个Excel的多个子表"""
    print("💾 正在生成Excel文件...")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for df, name in zip(data_frames, sheet_names):
            df.to_excel(writer, sheet_name=name, index=False)
    print(f"✅ 全部处理完成！结果已保存至：{OUTPUT_FILE}")

def run_pipeline():
    """主流程运行函数"""
    print("=== 数据清洗自动化脚本启动 ===")
    try:
        df_raw = load_data(INPUT_FILE)
        df_dedup = remove_duplicates(df_raw)
        df_filled = fill_missing_values(df_dedup)
        df_grouped = group_aggregate(df_filled)
        export_to_excel(
            [df_raw, df_dedup, df_filled, df_grouped],
            ['原始数据', '按姓名去重', '填充空值后', '按部门聚合']
        )
    except FileNotFoundError:
        print(f"❌ 错误：找不到文件 '{INPUT_FILE}'，请确认文件在同级目录下。")
    except Exception as e:
        print(f"❌ 发生未知错误：{e}")

def show_menu():
    """交互式菜单"""
    while True:
        print("=" * 30)
        print("请选择操作：")
        print("1. 运行完整数据清洗流程")
        print("2. 仅查看各部门销售聚合结果")
        print("3. 退出")
        print("=" * 30)
        choice = input("输入数字选择：")
        if choice == '1':
            run_pipeline()
        elif choice == '2':
            try:
                df = load_data(INPUT_FILE)
                df = remove_duplicates(df)
                df = fill_missing_values(df)
                df_grouped = group_aggregate(df)
                print("(仅预览，未生成Excel文件)")
            except Exception as e:
                print(f"❌ 错误：{e}")
        elif choice == '3':
            print("👋 脚本已退出，加油！")
            break
        else:
            print("⚠️ 无效输入，请重新输入数字。")

if __name__ == "__main__":
    show_menu()