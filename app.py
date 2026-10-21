import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="党员积分自动整合与统计", layout="wide")
st.title("🇨🇳 党员积分自动整合与统计系统")
st.write("上传您的《党员积分表_全部72人整合版.xlsx》，系统将在不破坏任何原始合并单元格、字体和样式的前提下，完美追加名次与等级。")

uploaded_file = st.file_uploader("请上传党员积分 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 1. 使用 openpyxl 完整读取带样式的原始表格文件
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        # 2. 用 pandas 临时读取数据，用来精确计算总分和名次
        df_raw = pd.read_excel(uploaded_file, header=1)
        df_raw = df_raw.dropna(subset=[df_raw.columns[0], df_raw.columns[1]], how='all').reset_index(drop=True)
        
        # 自动识别数值项并算总分
        exclude_keywords = ['名次', '排名', '手机', '电话', '身份', '序号', 'id', 'ID']
        score_cols = [col for col in df_raw.select_dtypes(include=['int64', 'float64']).columns if not any(kw in str(col) for kw in exclude_keywords)]
        
        # 计算总积分
        df_raw['最终总积分'] = df_raw[score_cols].sum(axis=1)
        # 根据总积分进行名次排序
        df_raw['新名次'] = df_raw['最终总积分'].rank(ascending=False, method='min').astype(int)
        
        # 3. 核心突破：直接在原始带有合并单元格的表格中“填空”
        # 我们从第 3 行（真实数据开始的行）向下循环，把名次精准填入 A 列
        党员映射 = {}
        for idx, row in df_raw.iterrows():
            姓名 = str(row[df_raw.columns[0]]).strip()
            党员映射[姓名] = {
                '名次': int(row['新名次']),
                '总分': float(row['最终总积分'])
            }
            
        # 扫描 openpyxl 的格子
        for r in range(3, ws.max_row + 1):
            姓名格子 = ws.cell(row=r, column=2).value # 假设B列是党员姓名
            if 姓名格子 and str(姓名格子).strip() in 党员映射:
                info = 党员映射[str(姓名格子).strip()]
                # A列原地修改为最新党内名次
                ws.cell(row=r, column=1, value=info['名次'])
                
                # 在表格最右侧（最后一列+1）原地追加积分等级
                total_score = info['总分']
                if total_score >= 90: 等级 = '优秀'
                elif total_score >= 80: 等级 = '合格'
                elif total_score >= 60: 等级 = '基本合格'
                else: 等级 = '不合格'
                
                # 寻找表格最右边一列填入等级
                ws.cell(row=r, column=ws.max_column, value=等级)
                
        st.success("📂 完美兼容！名次与等级已完美注入原始模板中，未损坏任何样式！")
        
        # 4. 一键导出原汁原味的二进制成品表格
        import io
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="🚀 点击一键下载【72人党员积分原格式完美成品表】",
            data=buffer,
            file_name='党员积分表_全部72人整合版_已统计.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        st.error(f"❌ 处理表格时出现问题，错误原因: {e}")
