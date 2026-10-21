import streamlit as st
import pandas as pd
import io
import openpyxl

# 1. 网页标签与配置
st.set_page_config(page_title="党员积分卡一键自动分装生成系统", layout="wide")

# 2. 极致精简的商业化视觉：彻底删除一切额外杂质字眼，只保留核心系统名称
st.title("党员积分卡一键自动分装生成系统")
st.caption("🚀 专为基层党组织定制的办公自动化工具")

# 3. 直观且规范的【使用说明】
st.markdown("""
---
### 📖 使用说明
为了确保系统能够精准完成数据自动誊抄，请确保您上传的两个 Excel 文件满足以下基础格式：

* **📁 1. 全员数据汇总表（Excel）**
  * 表格的 **第 2 行** 必须为真正的列标题（如：名次、党员姓名、党支部、党小组等）。
  * 表格的 **第二列** 必须包含党员的名字。
  * 从 **第四列** 开始，依次往后排列每一项考核的得分（自评分、考核分交替排列）。

* **📄 2. 标准考核模板（Excel）**
  * 表格的 **B 列（第二列）** 必须为具体的考评指标内容，且文本中带有 `1. ` `2. ` `3. ` 等数字序号。
  * 表格的 **C 列和 D 列（第三、四列）** 分别留空，用于系统自动填入自评得分和考核得分。
---
""")

# 4. 双文件上传组件
col1, col2 = st.columns(2)
with col1:
    data_file = st.file_uploader("📥 第一步：请上传【全员数据汇总表】（所有人员得分汇总）", type=["xlsx", "xls"])
with col2:
    template_file = st.file_uploader("📥 第二步：请上传【标准空白考核模板】（各车间通用的空白格式）", type=["xlsx", "xls"])

# 5. 核心逻辑控制
if data_file is not None and template_file is not None:
    # 增加手动点击生成的控制状态，防止自动刷新导致反复计算
    if "generated_data" not in st.session_state:
        st.session_state.generated_data = None

    st.markdown("<br>", unsafe_allow_html=True)
    # 显式提供【开始生成】按钮
    if st.button("🔥 开始一键生成党员积分卡", type="primary"):
        with st.spinner("🔄 正在精密对齐数据并自动分装，请稍候..."):
            try:
                # 读取汇总数据
                data_engine = 'openpyxl' if data_file.name.endswith('xlsx') else 'xlrd'
                df_data = pd.read_excel(data_file, header=1, engine=data_engine)
                
                # 智能在所有列名中寻找包含“姓名”或“党员”的列
                name_col = None
                for col in df_data.columns:
                    if '姓名' in str(col) or '党员' in str(col):
                        name_col = col
                        break
                if name_col is None:
                    name_col = df_data.columns
                
                # 提取所有的分数项列
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
                        
                    # 创建以党员姓名命名的独立工作表
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
                    
                    # 纯净填入数据并进行对齐格式统一
                    score_index = 0
                    for r in range(3, ws_new.max_row + 1):
                        val_b = str(ws_new.cell(row=r, column=2).value)
                        if val_b and any(f"{i}." in val_b for i in range(1, 10)):
                            
                            # 强行指定数值在格子内的对齐方式（水平居中、垂直居中），并保留一位小数格式
                            center_alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                            
                            # 填入自评得分（C列）
                            if score_index < len(score_cols):
                                cell_c = ws_new.cell(row=r, column=3, value=row[score_cols[score_index]])
                                cell_c.alignment = center_alignment
                                cell_c.number_format = '0.0'
                                score_index += 1
                                
                            # 填入考核得分（D列）
                            if score_index < len(score_cols):
                                cell_d = ws_new.cell(row=r, column=4, value=row[score_cols[score_index]])
                                cell_d.alignment = center_alignment
                                cell_d.number_format = '0.0'
                                score_index += 1
                    
                    generated_count += 1

                # 将结果存入二进制流
                buffer = io.BytesIO()
                output_wb.save(buffer)
                buffer.seek(0)
                
                st.session_state.generated_data = buffer
                st.success(f"🎉 分装成功！已为您自动生成了全员独立积分卡表格。")
                
            except Exception as e:
                st.error(f"❌ 自动化分装时出现错误: {e}")

    # 6. 展示下载按钮
    if st.session_state.generated_data is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 点击一键下载【全员原样样式积分卡总表.xlsx】",
            data=st.session_state.generated_data,
            file_name='全员党员积分卡分装总表.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
