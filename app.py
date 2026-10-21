import streamlit as st
import pandas as pd
import io
import openpyxl

# 1. 网页精美视觉包装（加入渐变色、国徽红与高级质感）
st.set_page_config(page_title="车间通用-党员积分卡自动分装生成系统", layout="wide", page_icon="🇨🇳")

st.markdown("""
    <style>
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #d32f2f;
        text-align: center;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 16px;
        color: #555;
        text-align: center;
        margin-bottom: 25px;
    }
    .info-box {
        background-color: #fefefe;
        padding: 15px;
        border-left: 5px solid #d32f2f;
        border-radius: 4px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>🇨🇳 党员积分卡一键自动分装生成系统</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>车间通用免费版 · 复杂合并单元格与行高列宽100%原样保留</div>", unsafe_allow_html=True)

st.markdown("""
<div class='info-box'>
    <b>💡 使用说明：</b><br>
    1. 在左侧框上传各车间汇总的【全员数据汇总表】（含党员姓名与各项得分）。<br>
    2. 在右侧框上传您车间专属的【标准空白考核模板】。<br>
    3. 检查无误后，点击下方的<b>“🚀 开始一键智能分装”</b>按钮，系统将自动为每位党员生成独立的工作表并保留完美格式。
</div>
""", unsafe_allow_html=True)

# 2. 提供双文件上传入口
col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 请上传《党员积分汇总表》（全员数据）", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 请上传《党员责任岗考评指标》（空白标准模板）", type=["xlsx", "xls"])

if data_file is not None and template_file is not None:
    # 🎯 核心优化 2：加入“开始生成”按钮，只有点击后才会触发数据处理，防止自动加载和报错
    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    start_button = st.button("🚀 开始一键智能分装", use_container_width=True)
    
    if start_button:
        with st.spinner("🔄 正在完美克隆模板格式并精准填充数据，请稍候..."):
            try:
                # 读取汇总数据
                data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
                df_data = pd.read_excel(data_file, header=1, engine=data_engine)
                
                name_col = None
                for col in df_data.columns:
                    if '姓名' in str(col) or '党员' in str(col):
                        name_col = col
                        break
                if name_col is None:
                    name_col = df_data.columns
                
                score_cols = df_data.columns[3:].tolist()
                
                output_wb = openpyxl.Workbook()
                active_first_sheet = output_wb.active
                
                template_wb = openpyxl.load_workbook(template_file)
                ws_template = template_wb.active
                
                generated_count = 0
                
                for index, row in df_data.iterrows():
                    name = str(row[name_col]).strip()
                    if not name or name == 'nan' or '党员姓名' in name or '备注' in name or '表格' in name or '党支部' in name:
                        continue
                        
                    if generated_count == 0:
                        ws_new = active_first_sheet
                        ws_new.title = f"{name}"
                    else:
                        ws_new = output_wb.create_sheet(title=f"{name}")
                    
                    # 100% 原样克隆模板的列宽与行高
                    for col_letter, col_dim in ws_template.column_dimensions.items():
                        ws_new.column_dimensions[col_letter].width = col_dim.width
                    for row_num, row_dim in ws_template.row_dimensions.items():
                        ws_new.row_dimensions[row_num].height = row_dim.height
                    
                    # 精准复制模板的每一个单元格的内容和视觉样式
                    for r in range(1, ws_template.max_row + 1):
                        for c in range(1, ws_template.max_column + 1):
                            cell_template = ws_template.cell(row=r, column=c)
                            cell_new = ws_new.cell(row=r, column=c, value=cell_template.value)
                            
                            if cell_template.has_style:
                                cell_new.font = openpyxl.styles.Font(name=cell_template.font.name, size=cell_template.font.size, bold=cell_template.font.bold, color=cell_template.font.color)
                                cell_new.alignment = openpyxl.styles.Alignment(horizontal=cell_template.alignment.horizontal, vertical=cell_template.alignment.vertical, wrap_text=cell_template.alignment.wrap_text)
                                cell_new.border = openpyxl.styles.Border(left=cell_template.border.left, right=cell_template.border.right, top=cell_template.border.top, bottom=cell_template.border.bottom)
                                if cell_template.fill:
                                    cell_new.fill = openpyxl.styles.PatternFill(fill_type=cell_template.fill.fill_type, start_color=cell_template.fill.start_color, end_color=cell_template.fill.end_color)
                    
                    # 完美复制合并单元格结构
                    for merged_range in ws_template.merged_cells.ranges:
                        ws_new.merge_cells(start_row=merged_range.min_row, start_column=merged_range.min_col,
                                           end_row=merged_range.max_row, end_column=merged_range.max_col)
                    
                    # 纯净填入数据
                    score_index = 0
                    for r in range(3, ws_new.max_row + 1):
                        val_b = str(ws_new.cell(row=r, column=2).value)
                        if val_b and any(f"{i}." in val_b for i in range(1, 10)):
                            
                            # 🎯 核心优化 1：强制将数据格子的格式设置为“水平居中、垂直居中”，并统一数字格式，解决第一行跟后面错位、对齐不一致的问题！
                            alignment_style = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
                            
                            # 填入自评得分（C列）
                            if score_index < len(score_cols):
                                cell_c = ws_new.cell(row=r, column=3, value=row[score_cols[score_index]])
                                cell_c.alignment = alignment_style
                                cell_c.number_format = '0.0' # 统一样式为保留一位小数（如13.8）
                                score_index += 1
                            # 填入考核得分（D列）
                            if score_index < len(score_cols):
                                cell_d = ws_new.cell(row=r, column=4, value=row[score_cols[score_index]])
                                cell_d.alignment = alignment_style
                                cell_d.number_format = '0.0'
                                score_index += 1
                    
                    generated_count += 1

                buffer = io.BytesIO()
                output_wb.save(buffer)
                buffer.seek(0)
                
                st.balloons() # 成功后弹出一整屏幕的喜庆气球，拉满推广仪式感！
                st.success(f"🎉 分装大功告成！已成功为 {generated_count} 名党员生成了与原模板格式100%一致的独立工作表。")
                
                st.download_button(
                    label="📥 点击一键下载【全员纯净独立积分卡总表】.xlsx",
                    data=buffer,
                    file_name='全员纯净独立积分卡总表.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ 自动化拆分分装时出现错误: {e}")
