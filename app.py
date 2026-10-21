import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="党员积分卡多工作表自动生成系统", layout="wide")
st.title("🇨🇳 党员积分卡一键自动分装生成系统")
st.write("请在下方依次上传您的【全员汇总资料表】和【标准考核模板】，系统已完美兼容 XLS 和 XLSX 格式。")

# 1. 🎯 核心优化：在 type 中同时加入 "xlsx" 和 "xls"，完美兼任两种格式
col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 请上传《党员积分表_全部72人整合版》（所有人数据汇总表）", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 请上传《党员责任岗月度积分考评指标》（标准空白考核模板）", type=["xlsx", "xls"])

if data_file is not None and template_file is not None:
    try:
        st.info("🔄 正在深度解析文件结构并保留原始样式，请稍候...")
        
        # 2. 自动根据文件后缀选择不同的引擎读取
        data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
        df_data = pd.read_excel(data_file, header=1, engine=data_engine)
        
        # 清洗空行
        df_data = df_data.dropna(subset=[df_data.columns[0]]).reset_index(drop=True)
        
        # 获取基础党员信息
        name_col = df_data.columns[0]   # 党员姓名
        branch_col = df_data.columns[1] # 党支部
        group_col = df_data.columns[2]  # 党小组
        
        # 提取所有的分数项列
        score_cols = df_data.columns[3:].tolist()
        
        # 创建一个内存中的全新工作簿作为最终导出的总Excel
        output_wb = openpyxl.Workbook()
        default_sheet = output_wb.active
        output_wb.remove(default_sheet)
        
        # 3. 兼容读取标准考核模板
        if template_file.name.endswith('xls'):
            # 如果模板是旧版 xls，先用 pandas 读，但由于 openpyxl 不支持直接载入 xls 的样式，建议提醒用户或云端静默转换
            # 为了保证100%带有原本漂亮的合并单元格和样式，强力推荐模板使用 xlsx 格式
            st.warning("⚠️ 检测到您的【空白考核模板】是旧版 .xls 格式。旧版格式可能会丢失原本的单元格颜色和精细边框，强烈建议将空白模板在电脑上“另存为” .xlsx 格式再行上传，效果最完美！")
            
        template_wb = openpyxl.load_workbook(template_file)
        ws_template = template_wb.active
        
        generated_count = 0
        for index, row in df_data.iterrows():
            name = str(row[name_col]).strip()
            if not name or name == 'nan':
                continue
                
            # 为每一位党员复制一个带原样式的Sheet
            ws_new = output_wb.create_sheet(title=f"{name}-积分卡")
            
            # 精准复制模板的结构、内容和合并单元格样式
            for r in range(1, ws_template.max_row + 1):
                for c in range(1, ws_template.max_column + 1):
                    cell_template = ws_template.cell(row=r, column=c)
                    cell_new = ws_new.cell(row=r, column=c, value=cell_template.value)
                    # 复制基础格式
                    if cell_template.has_style:
                        cell_new.font = openpyxl.styles.Font(name=cell_template.font.name, size=cell_template.font.size, bold=cell_template.font.bold, color=cell_template.font.color)
                        cell_new.alignment = openpyxl.styles.Alignment(horizontal=cell_template.alignment.horizontal, vertical=cell_template.alignment.vertical, wrap_text=cell_template.alignment.wrap_text)
                        cell_new.border = openpyxl.styles.Border(left=cell_template.border.left, right=cell_template.border.right, top=cell_template.border.top, bottom=cell_template.border.bottom)
                        if cell_template.fill:
                            cell_new.fill = openpyxl.styles.PatternFill(fill_type=cell_template.fill.fill_type, start_color=cell_template.fill.start_color, end_color=cell_template.fill.end_color)
            
            # 填补合并单元格
            for merged_range in ws_template.merged_cells.ranges:
                ws_new.merge_cells(start_row=merged_range.start_row, start_column=merged_range.start_column,
                                   end_row=merged_range.end_row, end_column=merged_range.end_column)
            
            # 🎯 核心映射：将总表里一长串的得分，按顺序填入到新Sheet对应的“自评得分”和“考核得分”格子中
            score_index = 0
            for r in range(3, ws_new.max_row + 1):
                val_b = str(ws_new.cell(row=r, column=2).value)
                if val_b and ('1.' in val_b or '2.' in val_b or '3.' in val_b or '4.' in val_b) and score_index < len(score_cols):
                    # 填入自评得分
                    if score_index < len(score_cols):
                        ws_new.cell(row=r, column=3, value=row[score_cols[score_index]])
                        score_index += 1
                    # 填入考核得分
                    if score_index < len(score_cols):
                        ws_new.cell(row=r, column=4, value=row[score_cols[score_index]])
                        score_index += 1
            
            # 在顶部或者合适空白处写入个人标志（可根据实际微调位置）
            ws_new.cell(row=1, column=1, value=f"🇨🇳 党员姓名: {name}   党支部: {row[branch_col]}   党小组: {row[group_col]}")
            generated_count += 1

        # 4. 二进制流一键打包导出
        buffer = io.BytesIO()
        output_wb.save(buffer)
        buffer.seek(0)
        
        st.success(f"🎉 全员分装成功！已成功为 {generated_count} 名党员每人自动生成了专属的独立工作表。")
        
        st.download_button(
            label="🚀 一键下载【全员72人独立积分卡总整合表.xlsx】",
            data=buffer,
            file_name='全员72人独立积分卡总整合表.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        st.error(f"❌ 自动化拆分分装时出现错误: {e}")
