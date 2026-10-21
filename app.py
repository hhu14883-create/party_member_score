import streamlit as st
import pandas as pd
import io
import openpyxl
import re
import copy

# 1. 网页标题
st.set_page_config(page_title="党员积分卡一键自动分装生成系统", layout="wide")

st.title("党员积分卡一键自动分装生成系统")
st.write("---") # 优雅的分割线

# 2. 左右并排的两大核心上传框
col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 上传：全员数据汇总表", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 上传：标准空白考核模板", type=["xlsx", "xls"])

st.write("<br>", unsafe_allow_html=True)

# 初始化内存状态
if "generated_data" not in st.session_state:
    st.session_state.generated_data = None

# 3. 创建居中按钮
b_col1, b_col2, b_col3 = st.columns([1, 2, 1])
with b_col2:
    start_btn = st.button("🔥 开始一键生成党员积分卡", type="primary", use_container_width=True)

# 4. 点击生成按钮后的核心业务逻辑
if start_btn:
    if data_file is None or template_file is None:
        st.error("❌ 请先在上方把两个 Excel 文件全部上传完整！")
    else:
        with st.spinner("🔄 正在精密对齐数据并自动分装，请稍候..."):
            try:
                # 读取汇总数据
                data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
                df_data = pd.read_excel(data_file, header=1, engine=data_engine)
                
                # 智能寻找包含“姓名”或“党员”的列
                name_col = None
                for col in df_data.columns:
                    if '姓名' in str(col) or '党员' in str(col):
                        name_col = col
                        break
                if name_col is None:
                    name_col = df_data.columns[1]
                
                # 提取分数项
                score_cols = df_data.columns[3:].tolist()
                
                # 创建总工作簿
                output_wb = openpyxl.Workbook()
                active_first_sheet = output_wb.active
                
                # 读取标准模板
                template_wb = openpyxl.load_workbook(template_file)
                ws_template = template_wb.active
                
                generated_count = 0
                
                for index, row in df_data.iterrows():
                    raw_name = str(row[name_col]).strip()
                    if not raw_name or raw_name == 'nan' or '党员姓名' in raw_name or '备注' in raw_name or '表格' in raw_name or '党支部' in raw_name:
                        continue
                        
                    # 清理 Excel Sheet 非法字符并防止重名
                    safe_name = re.sub(r'[\\/*?:\[\]\s]', '', raw_name)[:25]
                    sheet_title = safe_name
                    count = 1
                    while sheet_title in output_wb.sheetnames:
                        sheet_title = f"{safe_name}_{count}"
                        count += 1
                        
                    # 区分首 sheet 与后续 sheet
                    if generated_count == 0:
                        ws_new = active_first_sheet
                        ws_new.title = sheet_title
                    else:
                        ws_new = output_wb.create_sheet(title=sheet_title)
                    
                    # 克隆列宽与行高
                    for col_letter, col_dim in ws_template.column_dimensions.items():
                        ws_new.column_dimensions[col_letter].width = col_dim.width
                    for row_num, row_dim in ws_template.row_dimensions.items():
                        ws_new.row_dimensions[row_num].height = row_dim.height
                    
                    # 复制内容与基础格式
                    for r in range(1, ws_template.max_row + 1):
                        for c in range(1, ws_template.max_column + 1):
                            cell_template = ws_template.cell(row=r, column=c)
                            cell_new = ws_new.cell(row=r, column=c, value=cell_template.value)
                            
                            if cell_template.has_style:
                                if cell_template.font:
                                    cell_new.font = copy.copy(cell_template.font)
                                if cell_template.alignment:
                                    cell_new.alignment = copy.copy(cell_template.alignment)
                                if cell_template.border:
                                    cell_new.border = copy.copy(cell_template.border)
                                if cell_template.fill:
                                    cell_new.fill = copy.copy(cell_template.fill)
                    
                    # 复制合并单元格
                    for merged_range in ws_template.merged_cells.ranges:
                        ws_new.merge_cells(start_row=merged_range.min_row, start_column=merged_range.min_col,
                                           end_row=merged_range.max_row, end_column=merged_range.max_col)
                    
                    # 数据录入（使用正则精确匹配指标行，支持20条以上）
                    score_index = 0
                    for r in range(3, ws_new.max_row + 1):
                        val_b = str(ws_new.cell(row=r, column=2).value)
                        # 使用正则匹配类似 "1.", "2." 等开头的考评条目
                        if val_b and re.match(r'^\d+\.', val_b.strip()):
                            center_alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                            
                            # 填入自评（C列）
                            if score_index < len(score_cols):
                                cell_c = ws_new.cell(row=r, column=3, value=row[score_cols[score_index]])
                                cell_c.alignment = center_alignment
                                cell_c.number_format = '0.0'
                                score_index += 1
                                
                            # 填入考核（D列）
                            if score_index < len(score_cols):
                                cell_d = ws_new.cell(row=r, column=4, value=row[score_cols[score_index]])
                                cell_d.alignment = center_alignment
                                cell_d.number_format = '0.0'
                                score_index += 1
                    
                    generated_count += 1

                # 打包二进制流
                buffer = io.BytesIO()
                output_wb.save(buffer)
                buffer.seek(0)
                
                st.session_state.generated_data = buffer
                st.success(f"🎉 分装成功！已成功为您自动生成全部 {generated_count} 位党员的独立积分卡工作簿。")
                
            except Exception as e:
                st.error(f"❌ 分装失败，请检查表格，错误原因: {e}")

# 5. 生成成功后展示标准下载按钮
if st.session_state.generated_data is not None:
    st.write("<br>", unsafe_allow_html=True)
    d_col1, d_col2, d_col3 = st.columns([1, 2, 1])
    with d_col2:
        st.download_button(
            label="📥 点击一键下载【全员原样样式积分卡总表.xlsx】",
            data=st.session_state.generated_data,
            file_name='全员党员积分卡分装总表.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )
