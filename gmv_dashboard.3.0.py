import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from typing import Dict

# 全局周期配置
PERIOD_ORDER = ["5.1-5.7", "5.8-5.14", "5.15-5.21", "5.22-5.28"]
COLOR_PALETTE = px.colors.qualitative.Bold
PERIOD_COLOR_MAP = {
    "5.1-5.7": "#4472C4",
    "5.8-5.14": "#ED7D31",
    "5.15-5.21": "#A55194",
    "5.22-5.28": "#70AD47"
}

# 获取店铺名称
def get_store_from_filename(file: Path) -> str:
    return file.stem.split("_")[0].strip()

# 数据加载函数（云端相对路径，兼容Streamlit Cloud）
@st.cache_data(ttl=3600)
def load_all_excel(folder: str = "./各店铺周汇总报表") -> Dict[str, pd.DataFrame]:
    excel_dir = Path(folder)
    excel_files = list(excel_dir.glob("*_*周汇总报表.xlsx"))
    st.info(f"扫描到 {len(excel_files)} 个店铺Excel文件")

    all_dashboard = []
    all_sku_total = []
    all_sku_week = []
    all_gmv_detail = []

    for f in excel_files:
        store = get_store_from_filename(f)
        try:
            dash = pd.read_excel(f, sheet_name="Dashboard", engine="openpyxl")
            dash.columns = dash.columns.astype(str).str.strip()
            week = dash[["统计周期", "GMV", "销量"]].dropna().copy()
            week["店铺"] = store

            def get_val(label):
                match_col = dash.iloc[:, 0].str.contains(label, na=False)
                return dash.loc[match_col, "Unnamed: 1"].values[0]

            week["订单数"] = get_val("订单数")
            week["客单价"] = get_val("客单价")
            week["件单价"] = get_val("件单价")
            all_dashboard.append(week)
        except Exception:
            continue

        try:
            sku_total = pd.read_excel(f, sheet_name="SKU全周期汇总", engine="openpyxl")
            sku_total.columns = sku_total.columns.str.strip()
            sku_total["店铺"] = store
            all_sku_total.append(sku_total)
        except Exception:
            pass

        try:
            sku_week = pd.read_excel(f, sheet_name="SKU周度明细", engine="openpyxl")
            sku_week.columns = sku_week.columns.str.strip()
            sku_week["店铺"] = store
            all_sku_week.append(sku_week)
        except Exception:
            pass

        try:
            gmv_detail = pd.read_excel(f, sheet_name="GMV明细", engine="openpyxl")
            gmv_detail.columns = gmv_detail.columns.str.strip()
            gmv_detail["店铺"] = store
            all_gmv_detail.append(gmv_detail)
        except Exception:
            pass

    result = {}
    if all_dashboard:
        result["dashboard"] = pd.concat(all_dashboard, ignore_index=True)
        result["dashboard"]["统计周期"] = pd.Categorical(
            result["dashboard"]["统计周期"],
            categories=PERIOD_ORDER,
            ordered=True
        )
    if all_sku_total:
        result["sku_total"] = pd.concat(all_sku_total, ignore_index=True)
    if all_sku_week:
        result["sku_week"] = pd.concat(all_sku_week, ignore_index=True)
    if all_gmv_detail:
        result["gmv_detail"] = pd.concat(all_gmv_detail, ignore_index=True)
    return result

# 计算环比增长率
def add_growth(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.sort_values(["店铺", "统计周期"])
    df["GMV环比%"] = df.groupby("店铺")["GMV"].pct_change() * 100
    df["销量环比%"] = df.groupby("店铺")["销量"].pct_change() * 100
    return df

# 生成店铺周期透视表
def create_store_pivot_table(raw_df: pd.DataFrame) -> pd.DataFrame:
    gmv_pivot = raw_df.pivot(index="店铺", columns="统计周期", values="GMV").reset_index()
    gmv_col_list = ["店铺"] + [p for p in PERIOD_ORDER if p in gmv_pivot.columns]
    gmv_pivot = gmv_pivot[gmv_col_list]
    gmv_pivot.columns = ["店铺"] + [f"{p} GMV" for p in PERIOD_ORDER if p in gmv_pivot.columns]

    sales_pivot = raw_df.pivot(index="店铺", columns="统计周期", values="销量").reset_index()
    sales_col_list = ["店铺"] + [p for p in PERIOD_ORDER if p in sales_pivot.columns]
    sales_pivot = sales_pivot[sales_col_list]
    sales_pivot.columns = ["店铺"] + [f"{p} 销量" for p in PERIOD_ORDER if p in sales_pivot.columns]

    merge_df = pd.merge(gmv_pivot, sales_pivot, on="店铺", how="left")
    last_two_period = PERIOD_ORDER[-2:]
    prev_p, curr_p = last_two_period[0], last_two_period[1]
    merge_df["GMV环比变化额"] = merge_df[f"{curr_p} GMV"] - merge_df[f"{prev_p} GMV"]
    merge_df["GMV环比%"] = (merge_df[f"{curr_p} GMV"] / merge_df[f"{prev_p} GMV"]) * 100
    merge_df["销量环比%"] = ((merge_df[f"{curr_p} 销量"] - merge_df[f"{prev_p} 销量"]) / merge_df[f"{prev_p} 销量"]) * 100
    merge_df["排名"] = merge_df[f"{curr_p} GMV"].rank(ascending=False, method="min").astype(int)
    return merge_df

# 主页面程序
def main():
    st.set_page_config(page_title="多店铺GMV数据看板", layout="wide")
    st.title("📊 全店铺GMV&销量统一数据看板")
    st.markdown("---")

    with st.spinner("正在加载所有店铺数据..."):
        data = load_all_excel()

    if "dashboard" not in data or data["dashboard"].empty:
        st.error("❌ 未读取到有效店铺数据，请检查Excel文件是否上传到仓库")
        return

    dashboard = data["dashboard"]
    sku_total = data.get("sku_total", pd.DataFrame())
    sku_week = data.get("sku_week", pd.DataFrame())
    gmv_detail = data.get("gmv_detail", pd.DataFrame())
    dashboard = add_growth(dashboard)

    # 侧边筛选面板
    st.sidebar.header("🔍 筛选面板")
    store_list = sorted(dashboard["店铺"].unique())
    selected_stores = st.sidebar.multiselect("选择店铺", store_list, default=store_list)
    selected_periods = st.sidebar.multiselect("选择周期", PERIOD_ORDER, default=PERIOD_ORDER)
    top_sku_num = st.sidebar.slider("Top爆款SKU展示数量", min_value=5, max_value=50, value=20)

    # 筛选后基础数据集
    filter_condition = (dashboard["店铺"].isin(selected_stores)) & (dashboard["统计周期"].isin(selected_periods))
    df_filtered = dashboard[filter_condition].copy()
    df_pivot_table = create_store_pivot_table(df_filtered)

    # ====================== 第一模块：全局总览5列指标卡片（优化颜色显示） ======================
    st.header("🏁 全局总览指标")
    all_period_total_gmv = df_filtered["GMV"].sum()
    latest_cycle = PERIOD_ORDER[-1]
    prev_cycle = PERIOD_ORDER[-2]
    current_total_gmv = df_pivot_table[f"{latest_cycle} GMV"].sum()
    last_cycle_total_gmv = df_pivot_table[f"{prev_cycle} GMV"].sum()
    gmv_change_amount = current_total_gmv - last_cycle_total_gmv
    gmv_change_rate = (gmv_change_amount / last_cycle_total_gmv) * 100 if last_cycle_total_gmv != 0 else 0
    top_growth_store = df_pivot_table.sort_values("GMV环比%", ascending=False).iloc[0]["店铺"]

    # 5列布局，环比自动变色
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("全周期总GMV", f"{all_period_total_gmv:,.2f}")
    with col2:
        st.metric(f"本期({latest_cycle})GMV", f"{current_total_gmv:,.2f}")
    with col3:
        st.metric("GMV环比变化额", f"{gmv_change_amount:,.2f}")
    with col4:
        delta_text = f"{gmv_change_rate:.1f}%"
        # 负数红色，正数绿色
        delta_color = "inverse" if gmv_change_rate < 0 else "normal"
        st.metric("GMV环比变化率", f"{gmv_change_rate:.1f}%", delta=delta_text, delta_color=delta_color)
    with col5:
        st.metric("GMV增长TOP店铺", top_growth_store)

    # 新增业务状态提示文字
    st.markdown("### 📌 业务状态提示")
    if gmv_change_rate < -10:
        st.error(f"⚠️ 本周GMV环比大幅下滑{abs(gmv_change_rate):.1f}%，请核查各店铺销量、广告投放数据")
    elif gmv_change_rate < 0:
        st.info(f"ℹ️ 本周GMV小幅回落{abs(gmv_change_rate):.1f}%，整体波动平稳，可重点参考增长店铺【{top_growth_store}】运营策略")
    else:
        st.success(f"✅ 本周GMV环比上涨{gmv_change_rate:.1f}%，全店铺销售额提升")
    st.caption(f"💡 增长标杆店铺【{top_growth_store}】，可下拉页面查看TOP SKU榜单，提取爆款商品运营方案")
    st.markdown("---")

    # ====================== 第二模块：多周期店铺对比图表 ======================
    st.header("📈 多周期店铺对比图表")
    c_compare1, c_compare2, c_compare3 = st.columns(3)

    with c_compare1:
        fig_store_gmv = px.bar(
            df_filtered,
            x="店铺",
            y="GMV",
            color="统计周期",
            barmode="group",
            title="各店铺多周期GMV对比",
            color_discrete_map=PERIOD_COLOR_MAP
        )
        fig_store_gmv.update_layout(height=400)
        st.plotly_chart(fig_store_gmv, use_container_width=True)

    with c_compare2:
        fig_store_sales = px.bar(
            df_filtered,
            x="店铺",
            y="销量",
            color="统计周期",
            barmode="group",
            title="各店铺多周期销量对比",
            color_discrete_map=PERIOD_COLOR_MAP
        )
        fig_store_sales.update_layout(height=400)
        st.plotly_chart(fig_store_sales, use_container_width=True)

    with c_compare3:
        cycle_summary_list = []
        for p in PERIOD_ORDER:
            cycle_summary_list.append({
                "周期": p,
                "总GMV": df_pivot_table[f"{p} GMV"].sum(),
                "总销量": df_pivot_table[f"{p} 销量"].sum()
            })
        cycle_summary_df = pd.DataFrame(cycle_summary_list)
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Bar(
            x=cycle_summary_df["周期"],
            y=cycle_summary_df["总销量"],
            name="总销量",
            marker_color="#70AD47"
        ))
        fig_trend.add_trace(go.Line(
            x=cycle_summary_df["周期"],
            y=cycle_summary_df["总GMV"],
            name="总GMV",
            marker_color="#4472C4",
            yaxis="y2",
            mode="lines+markers"
        ))
        fig_trend.update_layout(
            title="全店铺周期销量(柱) & GMV(折线)",
            height=400,
            yaxis=dict(title="销量"),
            yaxis2=dict(title="GMV", overlaying="y", side="right")
        )
        st.plotly_chart(fig_trend, use_container_width=True)
    st.markdown("---")

    # ====================== 第三模块：单店铺筛选周度趋势 ======================
    st.header("📉 单店铺筛选周度趋势")
    c_single1, c_single2 = st.columns(2)
    with c_single1:
        fig_single_gmv = px.line(
            df_filtered,
            x="统计周期",
            y="GMV",
            color="店铺",
            markers=True,
            title="筛选店铺周GMV变化趋势",
            color_discrete_sequence=COLOR_PALETTE
        )
        fig_single_gmv.update_layout(height=400)
        st.plotly_chart(fig_single_gmv, use_container_width=True)
    with c_single2:
        fig_single_sales = px.bar(
            df_filtered,
            x="统计周期",
            y="销量",
            color="店铺",
            barmode="group",
            title="筛选店铺周销量对比",
            color_discrete_sequence=COLOR_PALETTE
        )
        fig_single_sales.update_layout(height=400)
        st.plotly_chart(fig_single_sales, use_container_width=True)
    st.markdown("---")

    # ====================== 第四模块：各店铺GMV环比变化率横向柱状图 ======================
    st.header("📊 各店铺GMV环比变化率")
    fig_growth_bar = px.bar(
        df_pivot_table,
        y="店铺",
        x="GMV环比%",
        orientation="h",
        title="各店铺GMV环比变化率",
        color="GMV环比%",
        color_continuous_scale=["#ff4d4f", "#4472C4"]
    )
    fig_growth_bar.update_layout(height=450)
    st.plotly_chart(fig_growth_bar, use_container_width=True)
    st.markdown("---")

    # ====================== 第五模块：各店铺全周期GMV市场占比 ======================
    st.header("🥧 各店铺全周期GMV市场占比")
    df_pivot_table["全周期累计GMV"] = df_pivot_table[[f"{p} GMV" for p in PERIOD_ORDER]].sum(axis=1)
    pie_fig = px.pie(
        df_pivot_table,
        values="全周期累计GMV",
        names="店铺",
        hole=0.2,
        title="店铺GMV占比分布",
        color_discrete_sequence=COLOR_PALETTE
    )
    pie_fig.update_traces(
        texttemplate="%{label}<br>%{value:,.2f}<br>占比%{percent:.1%}",
        textposition="inside"
    )
    pie_fig.update_layout(height=450)
    st.plotly_chart(pie_fig, use_container_width=True)
    st.markdown("---")

    # ====================== 第六模块：店铺本期GMV环比排名表格 ======================
    st.header("📈 店铺本期GMV环比排名")
    rank_display_df = df_pivot_table[["排名", "店铺", f"{latest_cycle} GMV", "GMV环比%"]].copy()
    st.dataframe(
        rank_display_df.style.format({
            f"{latest_cycle} GMV": "{:,.2f}",
            "GMV环比%": "{:.1f}%"
        }),
        use_container_width=True,
        hide_index=True
    )
    st.markdown("---")

    # ====================== 第七模块：店铺4周期GMV&销量完整明细表格 ======================
    st.header("📋 店铺全周期GMV&销量明细")
    format_dict = {}
    for p in PERIOD_ORDER:
        format_dict[f"{p} GMV"] = "{:,.2f}"
    format_dict["GMV环比变化额"] = "{:,.2f}"
    format_dict["全周期累计GMV"] = "{:,.2f}"
    format_dict["GMV环比%"] = "{:.2f}%"
    format_dict["销量环比%"] = "{:.2f}%"
    for p in PERIOD_ORDER:
        format_dict[f"{p} 销量"] = "{:,.0f}"

    st.dataframe(
        df_pivot_table.style.format(format_dict),
        use_container_width=True,
        hide_index=True
    )
    st.markdown("---")

    # ====================== 第八模块：全平台TOP爆款SKU横向柱状图（你需要的核心图表） ======================
    st.header("🏆 全平台TOP核心SKU GMV榜单")
    if not sku_total.empty:
        sku_filter_df = sku_total[sku_total["店铺"].isin(selected_stores)].copy()
        sku_filter_df["全周期累计GMV"] = sku_filter_df["GMV"]
        # 新增排序下拉选择框
        sort_type = st.selectbox(
            "SKU榜单排序依据",
            options=["全周期累计GMV 降序", "全周期累计GMV 升序", "销量 降序", "销量 升序"],
            index=0
        )
        # 根据选择执行排序
        if sort_type == "全周期累计GMV 降序":
            sku_sorted = sku_filter_df.sort_values("全周期累计GMV", ascending=False)
        elif sort_type == "全周期累计GMV 升序":
            sku_sorted = sku_filter_df.sort_values("全周期累计GMV", ascending=True)
        elif sort_type == "销量 降序":
            sku_sorted = sku_filter_df.sort_values("销量", ascending=False)
        else:
            sku_sorted = sku_filter_df.sort_values("全周期累计GMV", ascending=True)

        top_sku_df = sku_sorted.head(top_sku_num)
        sku_bar_fig = px.bar(
            top_sku_df,
            x="全周期累计GMV",
            y="SKU",
            color="店铺",
            orientation="h",
            title=f"TOP{top_sku_num} SKU 全周期GMV",
            color_discrete_sequence=COLOR_PALETTE
        )
        sku_bar_fig.update_layout(height=500)
        st.plotly_chart(sku_bar_fig, use_container_width=True)
        # SKU明细表格
        st.dataframe(
            top_sku_df[["店铺", "SKU", "商品名称", "全周期累计GMV", "销量"]]
            .style.format({"全周期累计GMV": "{:,.2f}", "销量": "{:.0f}"}),
            use_container_width=True,
            hide_index=True
        )

if __name__ == "__main__":
    main()
