import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="党员积分自动整合与统计", layout="wide")
st.title("🇨🇳 党员积分自动整合与统计系统")
st.write("上传《党员积分表_全部72人整合版.xlsx》，自动清洗数据、计算总积分、生成党内名次与积分等级。")

uploaded_file = st.file_uploader("请上传党员积分 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 1. 读取原始数据
        df = pd.read_excel(uploaded_file)
        st.success("📂 原始数据导入成功！正在为您自动分析...")
        
        # 2. 自动寻找姓名/党员名称列
        name_col = None
        for col in df.columns:
            if '姓名' in str(col) or '党员' in str(col) or '名字' in str(col):
                name_col = col
                break
        if not name_col:
            name_col = df.columns[0] # 如果没找到，默认第一列为标识列
            
        # 3. 自动识别所有数值型的“积分项”并计算总分
        # 排除包含‘名次’、‘排名’、‘手机’、‘身份证’等无关的数值列
        exclude_keywords = ['名次', '排名', '手机', '电话', '身份', '序号', 'id', 'ID']
        score_cols = []
        for col in df.select_dtypes(include=['int64', 'float64']).columns:
            if not any(kw in str(col) for kw in exclude_keywords):
                score_cols = score_cols + [col]
                
        # 核心整合：自动将所有分项积分相加得到最新总积分
        total_col_name = '最终总积分'
        df[total_col_name] = df[score_cols].sum(axis=1)
        
        # 4. 核心排序：按最终总积分进行降序排列
        df_sorted = df.sort_values(by=total_col_name, ascending=False).reset_index(drop=True)
        
        # 5. 核心插入：在表格第一列插入【党内名次】
        df_sorted.insert(0, '党内名次', df_sorted.index + 1)
        
        # 6. 核心划分：根据总分自动划分【积分等级】
        # 优秀：>=90, 合格: 80-89, 基本合格: 60-79, 不合格: <60
        def calculate_level(score):
            if score >= 90: return '优秀'
            elif score >= 80: return '合格'
            elif score >= 60: return '基本合格'
            else: return '不合格'
            
        df_sorted['积分等级'] = df_sorted[total_col_name].apply(calculate_level)
        
        # 7. 看板数据统计
        st.subheader("📊 核心数据快报")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总党员人数", f"{len(df_sorted)} 人")
        c2.metric("党内最高分", f"{df_sorted[total_col_name].max()} 分")
        c3.metric("全员平均分", f"{df_sorted[total_col_name].mean():.1f} 分")
        c4.metric("不及格人数（红牌）", f"{len(df_sorted[df_sorted[total_col_name] < 60])} 人")
        
        # 8. 预览前 10 名
        st.subheader("🏆 积分红榜（前10名党员）")
        st.dataframe(df_sorted[['党内名次', name_col, total_col_name, '积分等级']].head(10), use_container_width=True)
        
        # 9. 一键导出成品
        st.subheader("📥 导出完美成品表格")
        
        # 转换为 csv 格式供直接下载
        csv_data = df_sorted.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="🚀 点击一键下载【72人党员积分自动统计完美成品表】",
            data=csv_data,
            file_name='72人党员积分自动统计完美成品表.csv',
            mime='text/csv'
        )
        
    except Exception as e:
        st.error(f"❌ 分析表格时出现问题，请确保表格格式正常。错误原因: {e}")
