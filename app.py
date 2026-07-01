import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="社保公积金看板", layout="wide")
st.title("📊 2025年全国社保公积金数据看板")

@st.cache_data
def load_data():
    df = pd.read_excel("社保公积金阶段性周期性整理报表_全地区版.xlsx", sheet_name="月度明细数据表", header=1)
    return df

df = load_data()
df.columns = df.columns.str.strip()
df.rename(columns={
    "所属城市": "city", "所属分公司": "branch", "统计月份": "month",
    "参保人数": "headcount", "缴费基数": "base",
    "社保合计-总金额": "social_total", "公积金合计-总金额": "fund_total",
    "社保+公积金合计-总金额": "grand_total"
}, inplace=True)
df["month_num"] = df["month"].str.extract(r"(\d+)月").astype(float)

st.sidebar.header("🔎 筛选")
cities = ["全部"] + sorted(df["city"].unique().tolist())
selected_city = st.sidebar.selectbox("选择城市", cities)
month_range = st.sidebar.slider("月份", 1, 12, (1, 12))

filtered_df = df[(df["month_num"] >= month_range[0]) & (df["month_num"] <= month_range[1])]
if selected_city != "全部":
    filtered_df = filtered_df[filtered_df["city"] == selected_city]

col1, col2, col3, col4 = st.columns(4)
col1.metric("👥 累计人次", f"{filtered_df['headcount'].sum():,.0f}")
col2.metric("🏥 社保总额", f"{filtered_df['social_total'].sum():,.2f}")
col3.metric("🏠 公积金总额", f"{filtered_df['fund_total'].sum():,.2f}")
col4.metric("💰 合计", f"{filtered_df['grand_total'].sum():,.2f}")

city_summary = filtered_df.groupby("city")["grand_total"].sum().reset_index().sort_values("grand_total", ascending=False)
fig1 = px.bar(city_summary, x="city", y="grand_total", title="各城市缴费总额")
st.plotly_chart(fig1, use_container_width=True)

if selected_city != "全部":
    city_trend = filtered_df[filtered_df["city"] == selected_city].sort_values("month_num")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=city_trend["month"], y=city_trend["social_total"], name="社保"))
    fig2.add_trace(go.Scatter(x=city_trend["month"], y=city_trend["fund_total"], name="公积金"))
    st.plotly_chart(fig2, use_container_width=True)

with st.expander("📋 查看明细"):
    st.dataframe(filtered_df[["city", "month", "headcount", "grand_total"]])
