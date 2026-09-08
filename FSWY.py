import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill
from io import BytesIO

st.title("🏆 PK榜单筛选工具")
st.write("上传你的Excel文件，设置分数条件，一键筛选并下载结果！")

# 1. 文件上传组件
uploaded_file = st.file_uploader("选择Excel文件", type=["xlsx"])

if uploaded_file is not None:
    # 2. 读取上传的文件
    df = pd.read_excel(uploaded_file, engine='openpyxl', skiprows=1)
    
    st.write("📊 原始数据预览：")
    st.dataframe(df)
    
    # 3. 输入条件（滑块）
    min_score = st.slider("设置最低分数阈值", 0.0, 20.0, 10.0, 0.1)
    st.write(f"当前阈值：{min_score}")
    
    # 4. 筛选按钮
    if st.button("开始筛选"):
        # 筛选左边分数 > 阈值的
        filtered = df[df['分数'] > min_score]
        
        st.write(f"✅ 筛选出 {len(filtered)} 条记录（分数 > {min_score}）：")
        st.dataframe(filtered)
        
        # 5. 提供下载
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            filtered.to_excel(writer, index=False, sheet_name='筛选结果')
        
        st.download_button(
            label="📥 下载筛选结果",
            data=output.getvalue(),
            file_name="筛选结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )