"""
党员积分自动整合与统计 Web APP
基于 Streamlit + pandas + openpyxl
"""

from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# 页面配置
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="党员积分自动整合与统计",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# 业务常量
# ---------------------------------------------------------------------------
EXPECTED_FILENAME = "党员积分表_全部72人整合版.xlsx"
EXPECTED_COUNT = 72

# 常见非积分数值列（识别为数值型但仍排除出积分项）
EXCLUDE_NUMERIC_KEYWORDS = (
    "序号",
    "编号",
    "工号",
    "学号",
    "身份证",
    "年龄",
    "出生",
    "电话",
    "手机",
    "邮编",
    "年份",
    "年度",
    "季度",
    "月份",
    "名次",
    "排名",
)

# 可能已存在的总分列名关键词
TOTAL_SCORE_KEYWORDS = ("总积分", "总分", "最终总积分", "合计", "汇总")

# 可能的姓名/基本信息列关键词（用于优先展示）
INFO_KEYWORDS = (
    "姓名",
    "名字",
    "党员",
    "支部",
    "党组织",
    "职务",
    "岗位",
    "部门",
    "单位",
    "性别",
    "民族",
)


# ---------------------------------------------------------------------------
# 数据处理函数
# ---------------------------------------------------------------------------
def _is_excluded_column(col_name: str) -> bool:
    """判断数值列是否应排除在积分项之外。"""
    name = str(col_name).strip()
    return any(kw in name for kw in EXCLUDE_NUMERIC_KEYWORDS)


def _find_total_score_column(df: pd.DataFrame) -> Optional[str]:
    """查找已有的总分列。"""
    for col in df.columns:
        name = str(col).strip()
        if any(kw == name or kw in name for kw in TOTAL_SCORE_KEYWORDS):
            if pd.api.types.is_numeric_dtype(df[col]):
                return col
    return None


def _detect_info_columns(df: pd.DataFrame) -> list[str]:
    """识别党员基本信息列（非数值或含信息关键词的列）。"""
    info_cols: list[str] = []
    for col in df.columns:
        name = str(col).strip()
        if any(kw in name for kw in INFO_KEYWORDS):
            info_cols.append(col)
        elif not pd.api.types.is_numeric_dtype(df[col]):
            # 非数值且尚未收录
            if col not in info_cols:
                info_cols.append(col)
    return info_cols


def _detect_score_columns(df: pd.DataFrame, total_col: Optional[str]) -> list[str]:
    """自动提取所有数值型积分项列。"""
    score_cols: list[str] = []
    for col in df.columns:
        if col == total_col:
            continue
        if not pd.api.types.is_numeric_dtype(df[col]):
            continue
        if _is_excluded_column(col):
            continue
        score_cols.append(col)
    return score_cols


def _grade_by_score(score: float) -> str:
    """按分段划分积分等级。"""
    if pd.isna(score):
        return "未知"
    if score >= 90:
        return "优秀"
    if score >= 80:
        return "合格"
    if score >= 60:
        return "基本合格"
    return "不合格"


def clean_and_enrich(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str], str]:
    """
    清洗数据并生成：最终总积分、党内名次、积分等级。

    Returns
    -------
    result_df : 处理后的完整表
    score_cols : 识别出的积分项列名
    total_col_name : 使用的总积分列名
    """
    work = df.copy()

    # 去除全空行、全空列
    work = work.dropna(how="all").dropna(axis=1, how="all")
    # 统一列名空白
    work.columns = [str(c).strip() for c in work.columns]

    # 尝试将“看起来像数字”的 object 列转为数值（清洗）
    for col in work.columns:
        if work[col].dtype == object:
            converted = pd.to_numeric(work[col], errors="coerce")
            # 若转换后有效数值比例较高，则视为数值列
            non_null_orig = work[col].notna().sum()
            if non_null_orig > 0 and converted.notna().sum() / non_null_orig >= 0.7:
                work[col] = converted

    existing_total = _find_total_score_column(work)
    score_cols = _detect_score_columns(work, existing_total)

    total_col_name = "最终总积分"
    if existing_total is not None:
        work[total_col_name] = pd.to_numeric(work[existing_total], errors="coerce")
    else:
        if score_cols:
            work[total_col_name] = work[score_cols].sum(axis=1, min_count=1)
        else:
            work[total_col_name] = pd.NA

    # 积分项缺失填 0（仅用于展示一致性，总分已按现有规则计算）
    for col in score_cols:
        work[col] = pd.to_numeric(work[col], errors="coerce")

    # 降序排名；同分并列（min 法：1,2,2,4）
    work["党内名次"] = (
        work[total_col_name]
        .rank(method="min", ascending=False)
        .astype("Int64")
    )
    work["积分等级"] = work[total_col_name].apply(_grade_by_score)

    # 按总积分降序、名次升序排列
    work = work.sort_values(
        by=[total_col_name, "党内名次"],
        ascending=[False, True],
        na_position="last",
    ).reset_index(drop=True)

    return work, score_cols, total_col_name


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """将 DataFrame 导出为 Excel 字节流，供下载。"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="党员积分统计")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
def main() -> None:
    st.title("党员积分自动整合与统计")
    st.caption(
        f"上传《{EXPECTED_FILENAME}》，自动清洗数据、计算总积分、生成党内名次与积分等级。"
    )

    uploaded = st.file_uploader(
        "请上传党员积分 Excel 文件",
        type=["xlsx", "xls"],
        help=f"建议文件名：{EXPECTED_FILENAME}",
    )

    if uploaded is None:
        st.info("👆 请先上传 Excel 文件以开始分析。")
        return

    if EXPECTED_FILENAME not in uploaded.name and uploaded.name != EXPECTED_FILENAME:
        st.warning(
            f"当前文件名为「{uploaded.name}」，"
            f"建议使用「{EXPECTED_FILENAME}」。系统仍将继续处理。"
        )

    try:
        raw_df = pd.read_excel(uploaded, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001
        st.error(f"读取 Excel 失败：{exc}")
        return

    if raw_df.empty:
        st.error("上传的表格为空，请检查文件内容。")
        return

    with st.spinner("正在清洗数据并生成统计结果…"):
        result_df, score_cols, total_col = clean_and_enrich(raw_df)

    info_cols = _detect_info_columns(result_df)
    # 展示列：基本信息 + 积分项 + 派生列（去重保序）
    display_order: list[str] = []
    for col in info_cols + score_cols + [total_col, "党内名次", "积分等级"]:
        if col in result_df.columns and col not in display_order:
            display_order.append(col)
    for col in result_df.columns:
        if col not in display_order:
            display_order.append(col)

    # ---------- 核心指标 ----------
    st.subheader("核心指标看板")
    scores = pd.to_numeric(result_df[total_col], errors="coerce")
    total_people = len(result_df)
    avg_score = float(scores.mean()) if scores.notna().any() else 0.0
    max_score = float(scores.max()) if scores.notna().any() else 0.0
    fail_count = int((scores < 60).sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总人数", f"{total_people} 人", delta=f"预期 {EXPECTED_COUNT} 人")
    c2.metric("平均分", f"{avg_score:.2f}")
    c3.metric("最高分", f"{max_score:.2f}")
    c4.metric("不及格人数", f"{fail_count} 人", help="总积分 < 60")

    if total_people != EXPECTED_COUNT:
        st.caption(f"提示：当前识别到 {total_people} 人，与预期 72 人不一致，请核对源表。")

    # ---------- 识别结果说明 ----------
    with st.expander("数据识别详情", expanded=False):
        st.write("**识别到的积分项列：**", score_cols if score_cols else "（未识别到数值型积分列）")
        st.write("**基本信息列：**", info_cols if info_cols else "（未识别）")
        st.write(f"**总积分列：** `{total_col}`")

    # ---------- 分布可视化 ----------
    st.subheader("积分分布可视化")
    chart_left, chart_right = st.columns(2)

    with chart_left:
        st.markdown("**积分等级人数分布**")
        grade_order = ["优秀", "合格", "基本合格", "不合格", "未知"]
        grade_counts = (
            result_df["积分等级"]
            .value_counts()
            .reindex(grade_order, fill_value=0)
        )
        st.bar_chart(grade_counts)

    with chart_right:
        st.markdown("**最终总积分分布（直方图）**")
        hist_df = scores.dropna().to_frame(name="最终总积分")
        if not hist_df.empty:
            # 分箱后用柱状图近似直方图
            bins = pd.cut(
                hist_df["最终总积分"],
                bins=[0, 60, 80, 90, 100, float("inf")],
                right=False,
                labels=["<60", "60-79", "80-89", "90-100", ">100"],
            )
            bin_counts = bins.value_counts().sort_index()
            st.bar_chart(bin_counts)
        else:
            st.write("暂无有效积分数据。")

    # ---------- 明细表 ----------
    st.subheader("党员积分明细（已按总积分降序）")
    st.dataframe(result_df[display_order], use_container_width=True, height=420)

    # ---------- 一键导出 ----------
    st.subheader("一键导出")
    export_df = result_df[display_order].copy()
    excel_bytes = to_excel_bytes(export_df)
    st.download_button(
        label="📥 下载自动生成的积分统计表（Excel）",
        data=excel_bytes,
        file_name="党员积分统计结果_含总积分名次等级.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
