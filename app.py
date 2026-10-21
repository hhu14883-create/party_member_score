import streamlit as st
import pandas as pd
import io
import openpyxl

# 1. 极致黑金美化主题配置
st.set_page_config(page_title="党员积分卡一键自动分装生成系统", layout="wide")

# 注入自定义精美 CSS 样式，彻底告别原生的简陋视觉
st.markdown("""
    <style>
    /* 隐藏顶部红线和多余页眉 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* 页面大背景及全局字体调整 */
    .stApp {
        background: linear-gradient(135deg, #0d1117 0%, #161b22 100%);
        color: #c9d1d9;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
    }
    
    /* 核心主标题美化 */
    h1 {
        color: #ffdd67 !important;
        text-align: center;
        font-weight: 800 !important;
        letter-spacing: 2px;
        padding-top: 30px;
        padding-bottom: 30px;
        text-shadow: 0px 4px 20px rgba(255, 221, 103, 0.2);
    }
    
    /* 上传卡片区域美化 */
    .uploadedFile {
        background-color: #21262d !important;
        border: 2px dashed #30363d !important;
        border-radius: 12px !important;
        padding: 20px !important;
    }
    
    /* 强行改变大按钮的视觉（醒目深蓝到亮蓝渐变、居中悬浮感） */
    div.stButton > button:first-child {
        background: linear-gradient(90deg, #1f6feb 0%, #58a6ff 100%) !important;
        color: white !important;
        font-size: 18px !important;
        font-weight: bold !important;
        padding: 14px 40px !important;
        border-radius: 30px !important;
        border: none !important;
        box-shadow: 0px 8px 24px rgba(31, 111, 235, 0.4) !important;
        transition: all 0.3s ease !important;
        display: block;
        margin: 0 auto !important;
        width: 60% !important;
    }
    
    div.stButton > button:first-child:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0px 12px 30px rgba(88, 166, 255, 0.6) !important;
    }
    
    /* 下载按钮美化（亮金渐变） */
    .download-btn-container div.stButton > button:first-child {
        background: linear-gradient(90deg, #d29922 0%, #f1e05a 100%) !important;
        color: #0d1117 !important;
        box-shadow: 0px 8px 24px rgba(210, 153, 34, 0.4) !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. 只有独立的主标题，其余副标题、说明全部彻底删除
st.markdown("<h1>党员积分卡一键自动分装生成系统</h1>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# 3. 左右并排的两大核心上传框（无任何文字干涉干扰）
col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 上传：全员数据汇总表", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 上传：标准空白考核模板", type=["xlsx", "xls"])

st.markdown("<br><br>", unsafe_allow_html=True)

# 初始化内存状态
if "generated_data" not in st.session_state:
    st.session_state.generated_data = None

# 4. 醒目放置一键生成大按钮
if st.button("🔥 开始一键生成党员积分卡", type="primary"):
    if data_file is None or template_file is None:
        st.error("❌ 请先在上方把两个 Excel 文件全部上传完整！")
    else:
        with st.spinner("🔄 正在精密对齐数据并自动分装，请稍候..."):
            try:
                # 读取汇总数据
                data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
                df_data = pd.read_excel(data_file, header=1, engine=data_engine)
                
                # 寻找包含“姓名”或“党员”的列
                name_col = None
                for col in df_data.columns:
                    if '姓名' in str(col) or '党员' in str(col):
                        name_col = col
                        break
                if name_col is None:
                    name_col = df_data.columns
                
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
                    name = str(row[name_col]).strip()
                    if not name or name == 'nan' or '党员姓名' in name or '备注' in name or '表格' in name or '党支部' in name:
                        continue
                        
                    # 区分首 sheet 与后续 sheet
                    if generated_count == 0:
                        ws_new = active_first_sheet
                        ws_new.title = f"{name}"
                    else:
                        ws_new = output_wb.create_sheet(title=f"{name}")
                    
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
                                cell_new.font = openpyxl.styles.Font(name=cell_template.font.name, size=cell_template.font.size, bold=cell_template.font.bold, color=cell_template.font.color)
                                cell_new.alignment = openpyxl.styles.Alignment(horizontal=cell_template.alignment.horizontal, vertical=cell_template.alignment.vertical, wrap_text=cell_template.alignment.wrap_text)
                                cell_new.border = openpyxl.styles.Border(left=cell_template.border.left, right=cell_template.border.right, top=cell_template.border.top, bottom=cell_template.border.bottom)
                                if cell_template.fill:
                                    cell_new.fill = openpyxl.styles.PatternFill(fill_type=cell_template.fill.fill_type, start_color=cell_template.fill.start_color, end_color=cell_template.fill.end_color)
                    
                    # 复制合并单元格
                    for merged_range in ws_template.merged_cells.ranges:
                        ws_new.merge_cells(start_row=merged_range.min_row, start_column=merged_range.min_col,
                                           end_row=merged_range.max_row, end_column=merged_range.max_col)
                    
                    # 数据录入
                    score_index = 0
                    for r in range(3, ws_new.max_row + 1):
                        val_b = str(ws_new.cell(row=r, column=2).value)
                        if val_b and any(f"{i}." in val_b for i in range(1, 10)):
                            
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

                # 打包
                buffer = io.BytesIO()
                output_wb.save(buffer)
                buffer.seek(0)
                
                st.session_state.generated_data = buffer
                st.success(f"🎉 处理完毕！已成功打包 72 名党员的纯净积分卡工作表。")
                
            except Exception as e:
                st.error(f"❌ 分装失败，请检查表格，错误原因: {e}")

# 5. 生成成功后，在最下方展示高亮金色下载按钮
if st.session_state.generated_data is not None:
    st.markdown("<br><div class='download-btn-container'>", unsafe_allow_html=True)
    st.download_button(
        label="📥 部署完成！点击下载【全员原样样式积分卡总表.xlsx】",
        data=st.session_state.generated_data,
        file_name='全员党员积分卡分装总表.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    st.markdown("</div>", unsafe_allow_html=True)
