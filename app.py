import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="通用数据看板", layout="wide")
st.title("📊 通用 Excel / CSV 数据看板")
st.markdown("上传你的报表文件，自由探索数据，无需改代码")

# ---------- 上传文件 ----------
uploaded_file = st.file_uploader("上传 Excel 或 CSV 文件", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # 读取文件
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            # Excel 文件可能含多个 sheet，让用户选
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            if len(sheets) > 1:
                sheet_name = st.selectbox("选择工作表", list(sheets.keys()))
                df = sheets[sheet_name]
            else:
                df = list(sheets.values())[0]
    except Exception as e:
        st.error(f"读取文件出错: {e}")
        st.stop()

    # 清掉表头前的空行（如果有）
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    st.success(f"✅ 加载成功！共 {len(df)} 行，{len(df.columns)} 列")

    # ---------- 侧边栏：数据概览 ----------
    with st.sidebar:
        st.subheader("🔍 数据筛选")

        # 自动识别数字列和文本列
        num_cols = df.select_dtypes(include=["number"]).columns.tolist()
        text_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

        # 文本列筛选
        for col in text_cols[:3]:  # 最多显示前3个文本列做筛选，避免太挤
            unique_vals = df[col].dropna().unique().tolist()
            if len(unique_vals) < 50:
                selected = st.multiselect(f"筛选 {col}", unique_vals, default=unique_vals)
                if selected:
                    df = df[df[col].isin(selected)]

        # 数字列范围筛选（显示前2个数字列）
        for col in num_cols[:2]:
            if not df[col].dropna().empty:
                min_val = float(df[col].min())
                max_val = float(df[col].max())
                if min_val < max_val:
                    range_val = st.slider(f"{col} 范围", min_val, max_val, (min_val, max_val))
                    df = df[(df[col] >= range_val[0]) & (df[col] <= range_val[1])]

        st.caption(f"当前行数: {len(df)}")

    # ---------- 指标卡片 ----------
    st.subheader("📈 关键指标")
    cols = st.columns(min(len(num_cols), 6))
    for i, col in enumerate(num_cols[:6]):
        if not df[col].dropna().empty:
            cols[i].metric(col, f"{df[col].sum():,.2f}")

    # ---------- 图表生成器 ----------
    st.subheader("📊 自由制图")

    col1, col2, col3 = st.columns(3)
    with col1:
        chart_type = st.selectbox("图表类型", ["柱状图", "折线图", "散点图", "饼图", "箱线图"])
    with col2:
        x_axis = st.selectbox("X轴 / 分类", [None] + df.columns.tolist())
    with col3:
        y_axis = st.selectbox("Y轴 / 数值", [None] + num_cols)

    if x_axis and y_axis:
        fig = None
        if chart_type == "柱状图":
            fig = px.bar(df, x=x_axis, y=y_axis, title=f"{y_axis} 按 {x_axis} 分布")
        elif chart_type == "折线图":
            fig = px.line(df, x=x_axis, y=y_axis, title=f"{y_axis} 按 {x_axis} 趋势")
        elif chart_type == "散点图":
            fig = px.scatter(df, x=x_axis, y=y_axis, title=f"{x_axis} vs {y_axis}")
        elif chart_type == "饼图":
            # 饼图需要聚合，按x分组求和y
            grouped = df.groupby(x_axis)[y_axis].sum().reset_index()
            fig = px.pie(grouped, names=x_axis, values=y_axis, title=f"{y_axis} 占比")
        elif chart_type == "箱线图":
            fig = px.box(df, x=x_axis, y=y_axis, title=f"{y_axis} 分布")

        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # ---------- 数据明细 ----------
    with st.expander("📋 查看全部数据"):
        st.dataframe(df, use_container_width=True, height=400)

    # ---------- 下载结果 ----------
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ 下载当前筛选后的数据 (CSV)", data=csv, file_name="filtered_data.csv", mime="text/csv")

else:
    st.info("👈 请先上传你的 Excel 或 CSV 文件")
