import streamlit as st
import pandas as pd
import io
import openpyxl
import re
import copy

# 1. 网页标题
st.set_page_config(page_title="党员积分卡一键自动分装生成系统", layout="wide")

st.title("党员积分卡一键自动分装生成系统（双向智能跳转）")
st.write("---")

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
    start_btn = st.button("🔥 开始一键生成党员积分卡（双向跳转）", type="primary", use_container_width=True)

# 4. 点击生成按钮后的核心业务逻辑
if start_btn:
    if data_file is None or template_file is None:
        st.error("❌ 请先在上方把两个 Excel 文件全部上传完整！")
    else:
        with st.spinner("🔄 正在精密生成双向跳转目录、对齐数据并自动分装，请稍候..."):
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
                
                # 读取标准模板
                template_wb = openpyxl.load_workbook(template_file)
                ws_template = template_wb.active
                
                # ---------------------------------------------------------
                # 第一步：创建“目录” Sheet (放在最前面)
                # ---------------------------------------------------------
                ws_catalog = output_wb.active
                ws_catalog.title = "目录"
                
                # 设置目录表头
                catalog_headers = ["姓名", "链接点击跳转栏", "总分"]
                thin_side = openpyxl.styles.Side(border_style="thin", color="000000")
                thin_border = openpyxl.styles.Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                
                for c_idx, h_text in enumerate(catalog_headers, start=1):
                    cell = ws_catalog.cell(row=1, column=c_idx, value=h_text)
                    cell.font = openpyxl.styles.Font(name="宋体", size=11, bold=True)
                    cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
                    cell.fill = openpyxl.styles.PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                
                ws_catalog.column_dimensions['A'].width = 18
                ws_catalog.column_dimensions['B'].width = 25
                ws_catalog.column_dimensions['C'].width = 15
                
                valid_rows = []
                for index, row in df_data.iterrows():
                    raw_name = str(row[name_col]).strip()
                    if not raw_name or raw_name == 'nan' or '党员姓名' in raw_name or '备注' in raw_name or '表格' in raw_name or '党支部' in raw_name:
                        continue
                    
                    # 清理 Sheet 非法字符并防止重名
                    safe_name = re.sub(r'[\\/*?:\[\]\s]', '', raw_name)[:25]
                    sheet_title = safe_name
                    count = 1
                    while sheet_title in output_wb.sheetnames or sheet_title == "目录":
                        sheet_title = f"{safe_name}_{count}"
                        count += 1
                        
                    valid_rows.append((row, sheet_title, raw_name))
                
                # ---------------------------------------------------------
                # 第二步：循环生成每个党员的个人 Sheet，并写入双向跳转链接
                # ---------------------------------------------------------
                for idx, (row, sheet_title, raw_name) in enumerate(valid_rows, start=2):
                    # 1. 写入目录行
                    cell_name = ws_catalog.cell(row=idx, column=1, value=raw_name)
                    cell_name.font = openpyxl.styles.Font(name="宋体", size=10)
                    cell_name.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    cell_name.border = thin_border
                    
                    # 目录页 -> 跳转个人表
                    cell_link = ws_catalog.cell(row=idx, column=2, value=f'=HYPERLINK("#\'{sheet_title}\'!A1", "点击跳转相应表格")')
                    cell_link.font = openpyxl.styles.Font(name="宋体", size=10, color="0000FF", underline="single")
                    cell_link.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    cell_link.border = thin_border
                    
                    # 计算总分
                    total_score = ""
                    try:
                        total_score = row.iloc[-1] if not pd.isna(row.iloc[-1]) else row.iloc[-2]
                    except:
                        total_score = 0.0
                        
                    cell_score = ws_catalog.cell(row=idx, column=3, value=total_score)
                    cell_score.font = openpyxl.styles.Font(name="宋体", size=10)
                    cell_score.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
                    cell_score.border = thin_border
                    cell_score.number_format = '0.0000'
                    
                    # 2. 创建党员个人 Sheet
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
                    
                    # ---------------------------------------------------------
                    # 3. 核心新增：把个人表格第一行的标题改成“点击标题返回目录”的超链接公式
                    # ---------------------------------------------------------
                    title_cell = ws_new.cell(row=1, column=1)
                    original_title_text = str(title_cell.value) if title_cell.value else "党员责任岗月度积分考评指标体系指导范本（作业岗位）"
                    # 清理可能存在的旧换行并加上提示
                    clean_title = original_title_text.split('\n')[0].strip()
                    display_title_text = f'{clean_title}\n（点击标题返回目录）'
                    
                    # 写入返回目录的 HYPERLINK 公式
                    title_cell.value = f'=HYPERLINK("#\'目录\'!A1", "{display_title_text}")'
                    title_cell.font = openpyxl.styles.Font(name="宋体", size=11, bold=True, color="0000FF")
                    title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)

                    # 数据录入
                    score_index = 0
                    for r in range(3, ws_new.max_row + 1):
                        val_b = str(ws_new.cell(row=r, column=2).value)
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

                # 打包二进制流
                buffer = io.BytesIO()
                output_wb.save(buffer)
                buffer.seek(0)
                
                st.session_state.generated_data = buffer
                st.success(f"🎉 分装成功！双向跳转（目录 ⇄ 个人表头）已全部配好，共生成 {len(valid_rows)} 位党员档案。")
                
            except Exception as e:
                st.error(f"❌ 分装失败，请检查表格，错误原因: {e}")

# 5. 生成成功后展示标准下载按钮
if st.session_state.generated_data is not None:
    st.write("<br>", unsafe_allow_html=True)
    d_col1, d_col2, d_col3 = st.columns([1, 2, 1])
    with d_col2:
        st.download_button(
            label="📥 点击一键下载【双向跳转党员积分卡总表.xlsx】",
            data=st.session_state.generated_data,
            file_name='党员积分卡分装总表_双向跳转.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )
