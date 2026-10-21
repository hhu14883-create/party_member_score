import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="党员积分卡多工作表自动生成系统", layout="wide")
st.title("🇨🇳 党员积分卡一键自动分装生成系统")
st.write("请在下方依次上传您的【全员汇总资料表】和【标准考核模板】，系统已完美兼容 XLS 和 XLSX 格式。")

col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 请上传《党员积分表_全部72人整合版》（所有人数据汇总表）", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 请上传《党员责任岗月度积分考评指标》（标准空白考核模板）", type=["xlsx", "xls"])

if data_file is not None and template_file is not None:
    try:
        st.info("🔄 正在深度解析文件结构并保留原始样式，请稍候...")
        
        # 自动根据文件后缀选择不同的引擎读取
        data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
        df_data = pd.read_excel(data_file, header=1, engine=data_engine)
        
        # 清洗空行
        df_data = df_data.dropna(subset=[df_data.columns]).reset_index(drop=True)
        
        # 获取党员姓名列
        name_col = df_data.columns
        
        # 提取所有的分数项列
        score_cols = df_data.columns[3:].tolist()
        
        # 创建一个内存中的全新工作簿作为最终导出的总Excel
        output_wb = openpyxl.Workbook()
        default_sheet = output_wb.active
        output_wb.remove(default_sheet)
        
        # 读取标准考核模板
        template_wb = openpyxl.load_workbook(template_file)
        ws_template = template_wb.active
        
        generated_count = 0
        for index, row in df_data.iterrows():
            name = str(row[name_col]).strip()
            if not name or name == 'nan' or '党员姓名' in name:
                continue
                
            # 🎯 核心优化 1：Sheet标签名称只要纯姓名，去掉所有多余后缀
            ws_new = output_wb.create_sheet(title=f"{name}")
            
            # 精准复制模板的结构、内容和样式
            for r in range(1, ws_template.max_row + 1):
                for c in range(1, ws_template.max_column + 1):
                    cell_template = ws_template.cell(row=r, column=c)
                    cell_new = ws_new.cell(row=r, column=c, value=cell_template.value)
                    # 完整保留原始的所有精美格式（字体、颜色、对齐、边框）
                    if cell_template.has_style:
                        cell_new.font = openpyxl.styles.Font(name=cell_template.font.name, size=cell_template.font.size, bold=cell_template.font.bold, color=cell_template.font.color)
                        cell_new.alignment = openpyxl.styles.Alignment(horizontal=cell_template.alignment.horizontal, vertical=cell_template.alignment.vertical, wrap_text=cell_template.alignment.wrap_text)
                        cell_new.border = openpyxl.styles.Border(left=cell_template.border.left, right=cell_template.border.right, top=cell_template.border.top, bottom=cell_template.border.bottom)
                        if cell_template.fill:
                            cell_new.fill = openpyxl.styles.PatternFill(fill_type=cell_template.fill.fill_type, start_color=cell_template.fill.start_color, end_color=cell_template.fill.end_color)
            
            # 完整复制合并单元格
            for merged_range in ws_template.merged_cells.ranges:
                ws_new.merge_cells(start_row=merged_range.min_row, start_column=merged_range.min_col,
                                   end_row=merged_range.max_row, end_column=merged_range.max_col)
            
            # 🎯 核心优化 2：彻底删除了之前强行在第1行写个人标志的乱套代码！
            # 老老实实只把数据往模板的对应格子里填，不更改模板原有的任何字样
            score_index = 0
            for r in range(3, ws_new.max_row + 1):
                val_b = str(ws_new.cell(row=r, column=2).value)
                if val_b and any(f"{i}." in val_b for i in range(1, 10)) and score_index < len(score_cols):
                    # 填入自评得分（C列）
                    if score_index < len(score_cols):
                        ws_new.cell(row=r, column=3, value=row[score_cols[score_index]])
                        score_index += 1
                    # 填入考核得分（D列）
                    if score_index < len(score_cols):
                        ws_new.cell(row=r, column=4, value=row[score_cols[score_index]])
                        score_index += 1
            
            generated_count += 1

        # 4. 二进制流一键打包导出
        buffer = io.BytesIO()
        output_wb.save(buffer)
        buffer.seek(0)
        
        st.success(f"🎉 纯净分装成功！已成功为 {generated_count} 名党员每人自动生成了纯净独立的专属工作表。")
        
        st.download_button(
            label="🚀 一键下载【全员72人纯净积分卡总整合表.xlsx】",
            data=buffer,
            file_name='全员72人纯净积分卡总整合表.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        st.error(f"❌ 自动化拆分分装时出现错误: {e}")
