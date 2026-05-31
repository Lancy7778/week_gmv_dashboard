import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from typing import Dict
import re
import numpy as np

# 全局周期 - 使用 Session State 来存储
if 'period_order' not in st.session_state:
    st.session_state.period_order = []

COLOR_PALETTE = px.colors.qualitative.Bold

# 页面配置
st.set_page_config(
    page_title="多店铺GMV看板",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    /* 隐藏默认header */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}

    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }

    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        padding: 20px 15px;
        text-align: center;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
        height: 100%;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 40px rgba(0,0,0,0.15);
    }
    .metric-card.blue {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
    }
    .metric-card.green {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    .metric-card.orange {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    }
    .metric-card.purple {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    }
    .metric-card.red {
        background: linear-gradient(135deg, #f5576c 0%, #f093fb 100%);
    }
    .metric-card.gray {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border: 1px solid rgba(0,0,0,0.08);
    }
    .metric-card.gray .metric-label {
        color: #495057;
    }
    .metric-card.gray .metric-value {
        color: #212529;
    }
    .metric-card.gray .metric-sub {
        color: #6c757d;
    }

    .metric-label {
        color: rgba(255,255,255,0.9);
        font-size: 0.85rem;
        font-weight: 500;
        letter-spacing: 1px;
        margin-bottom: 12px;
        text-transform: uppercase;
    }
    .metric-value {
        color: white;
        font-size: 1.8rem;
        font-weight: bold;
        margin: 8px 0;
        white-space: nowrap;
    }
    .metric-sub {
        color: rgba(255,255,255,0.7);
        font-size: 0.75rem;
        margin-top: 8px;
    }
    .delta-positive {
        color: #00ff88;
        font-weight: bold;
    }
    .delta-negative {
        color: #ff6b6b;
        font-weight: bold;
    }

    .section-title {
        font-size: 1.4rem;
        font-weight: bold;
        color: #1e3c72;
        margin: 20px 0 15px 0;
        padding-left: 12px;
        border-left: 4px solid #4facfe;
    }
</style>
""", unsafe_allow_html=True)


# 获取店铺名
def get_store_from_filename(file: Path) -> str:
    return file.stem.split("_")[0].strip()


# 安全转换数值
def safe_float(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value):
            return 0.0
        return float(value)
    try:
        return float(value)
    except:
        return 0.0


def safe_int(value):
    if pd.isna(value):
        return 0
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value):
            return 0
        return int(value)
    try:
        return int(float(value))
    except:
        return 0


def format_num(num):
    if pd.isna(num) or np.isnan(num):
        return "0.00"
    return f"{num:,.2f}"


def sort_period_key(period):
    nums = re.findall(r"\d+", str(period))
    if len(nums) >= 4:
        return (int(nums[0]), int(nums[1]), int(nums[2]), int(nums[3]))
    elif len(nums) >= 2:
        return (int(nums[0]), int(nums[1]), 0, 0)
    return (0, 0, 0, 0)


# 加载数据
@st.cache_data(ttl=3600)
def load_all_excel(
        folder: str = r"D:\儒易工作内容\每周汇报GMV\测试周GMV自动化流程\各店铺周汇总报表"
):
    excel_dir = Path(folder)
    excel_files = list(excel_dir.glob("*_多周汇总报表.xlsx"))

    if not excel_files:
        st.error(f"❌ 未找到Excel文件，请检查路径: {folder}")
        return None, []

    all_dashboard = []

    for f in excel_files:
        store = get_store_from_filename(f)
        try:
            dash = pd.read_excel(f, sheet_name="Dashboard", engine="openpyxl")
            dash.columns = dash.columns.astype(str).str.strip()

            # 查找周期列
            period_col = None
            possible_period_cols = ['统计周期', '周期', 'Period', 'period']
            for col in possible_period_cols:
                if col in dash.columns:
                    period_col = col
                    break

            if period_col is None:
                for col in dash.columns:
                    if '周期' in col or 'period' in col.lower():
                        period_col = col
                        break

            if period_col is None:
                continue

            # 提取GMV和销量列
            gmv_col = 'GMV' if 'GMV' in dash.columns else None
            sales_col = '销量' if '销量' in dash.columns else None

            if gmv_col is None or sales_col is None:
                continue

            # 创建周度数据
            for idx, row in dash.iterrows():
                period_val = row[period_col]
                if pd.isna(period_val):
                    continue

                gmv_val = safe_float(row[gmv_col]) if pd.notna(row[gmv_col]) else 0
                sales_val = safe_int(row[sales_col]) if pd.notna(row[sales_col]) else 0

                all_dashboard.append({
                    "统计周期": str(period_val).strip(),
                    "GMV": gmv_val,
                    "销量": sales_val,
                    "店铺": store
                })

        except Exception as e:
            continue

    if not all_dashboard:
        return None, []

    df = pd.DataFrame(all_dashboard)

    # 获取所有唯一的周期并排序
    all_periods = df["统计周期"].unique().tolist()
    periods = sorted(all_periods, key=sort_period_key)

    # 设置为分类变量
    df["统计周期"] = pd.Categorical(df["统计周期"], categories=periods, ordered=True)

    return df, periods


# 环比计算
def add_growth(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()
    df = df.sort_values(["店铺", "统计周期"])

    df["GMV"] = pd.to_numeric(df["GMV"], errors='coerce').fillna(0)
    df["销量"] = pd.to_numeric(df["销量"], errors='coerce').fillna(0)

    df["GMV环比%"] = df.groupby("店铺")["GMV"].pct_change() * 100
    df["销量环比%"] = df.groupby("店铺")["销量"].pct_change() * 100

    df["GMV环比%"] = df["GMV环比%"].replace([float('inf'), -float('inf')], 0).fillna(0)
    df["销量环比%"] = df["销量环比%"].replace([float('inf'), -float('inf')], 0).fillna(0)

    return df


# 透视表
def create_store_pivot_table(raw_df: pd.DataFrame, periods: list):
    if len(periods) < 2:
        return pd.DataFrame()

    if raw_df.empty:
        return pd.DataFrame()

    raw_df = raw_df.copy()
    raw_df["GMV"] = pd.to_numeric(raw_df["GMV"], errors='coerce').fillna(0)
    raw_df["销量"] = pd.to_numeric(raw_df["销量"], errors='coerce').fillna(0)

    try:
        gmv_pivot = raw_df.pivot_table(
            index="店铺",
            columns="统计周期",
            values="GMV",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        exist_periods = [p for p in periods if p in gmv_pivot.columns]
        if not exist_periods:
            return pd.DataFrame()

        gmv_cols = ["店铺"] + exist_periods
        gmv_pivot = gmv_pivot[gmv_cols]
        gmv_pivot.columns = ["店铺"] + [f"{p} GMV" for p in exist_periods]

        sales_pivot = raw_df.pivot_table(
            index="店铺",
            columns="统计周期",
            values="销量",
            aggfunc="sum",
            fill_value=0
        ).reset_index()
        sales_pivot = sales_pivot[["店铺"] + exist_periods]
        sales_pivot.columns = ["店铺"] + [f"{p} 销量" for p in exist_periods]

        merge_df = pd.merge(gmv_pivot, sales_pivot, on="店铺", how="left")

        if len(exist_periods) >= 2:
            prev_p, curr_p = exist_periods[-2], exist_periods[-1]

            prev_gmv = merge_df[f"{prev_p} GMV"].replace(0, np.nan)
            merge_df["GMV环比变化额"] = merge_df[f"{curr_p} GMV"] - merge_df[f"{prev_p} GMV"]
            merge_df["GMV环比%"] = ((merge_df[f"{curr_p} GMV"] / prev_gmv) - 1) * 100
            merge_df["销量环比%"] = ((merge_df[f"{curr_p} 销量"] / merge_df[f"{prev_p} 销量"].replace(0,
                                                                                                      np.nan)) - 1) * 100

            merge_df["GMV环比%"] = merge_df["GMV环比%"].fillna(0).replace([float('inf'), -float('inf')], 0)
            merge_df["销量环比%"] = merge_df["销量环比%"].fillna(0).replace([float('inf'), -float('inf')], 0)
            merge_df["GMV环比变化额"] = merge_df["GMV环比变化额"].fillna(0)

            curr_gmv = merge_df[f"{curr_p} GMV"].replace([float('inf'), -float('inf')], 0)
            merge_df["排名"] = curr_gmv.rank(ascending=False, method="min").fillna(999).astype(int)

        gmv_cols = [f"{p} GMV" for p in exist_periods]
        merge_df["全周期累计GMV"] = merge_df[gmv_cols].sum(axis=1).fillna(0)

        return merge_df
    except Exception as e:
        return pd.DataFrame()


# 主函数
def main():
    # 标题
    st.markdown("""
    <div style='text-align: center; margin-bottom: 20px;'>
        <h1 style='color: #1e3c72; font-size: 2.2rem; margin-bottom: 5px;'>📊 全店铺GMV&销量统一数据看板</h1>
        <p style='color: #666; font-size: 0.9rem;'>实时监控 | 多维度分析 | 智能洞察</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    with st.spinner("加载数据中..."):
        dashboard_df, periods = load_all_excel()

    if dashboard_df is None or dashboard_df.empty:
        st.error("❌ 未读取到有效店铺数据，请确保Excel文件已生成")
        st.info("请先运行: python auto2.0.py --mode once")
        return

    # 保存到 session state
    st.session_state.period_order = periods

    # 添加环比
    dashboard_df = add_growth(dashboard_df)

    # 侧边栏筛选
    with st.sidebar:
        st.markdown("### 🔍 筛选面板")
        st.markdown("---")

        # 获取所有店铺
        store_list = sorted(dashboard_df["店铺"].unique())

        # 店铺选择
        selected_stores = st.multiselect(
            "🏪 选择店铺",
            store_list,
            default=store_list
        )

        # 周期选择 - 使用 session_state
        if st.session_state.period_order:
            selected_periods = st.multiselect(
                "📅 选择周期",
                st.session_state.period_order,
                default=st.session_state.period_order
            )
        else:
            st.error("⚠️ 未找到周期数据，请检查Excel文件")
            selected_periods = []

        top_sku_num = st.slider("🏆 Top爆款SKU数量", 5, 50, 20)

        st.markdown("---")
        st.caption(f"✅ 加载成功: {len(store_list)} 个店铺, {len(periods)} 个周期")
        if periods:
            st.caption(f"📅 周期: {', '.join(periods)}")

    # 检查筛选条件
    if not selected_stores:
        st.warning("⚠️ 请至少选择一个店铺")
        return

    if not selected_periods:
        st.warning("⚠️ 请至少选择一个周期")
        return

    # 筛选数据
    filter_condition = (dashboard_df["店铺"].isin(selected_stores)) & (dashboard_df["统计周期"].isin(selected_periods))
    df_filtered = dashboard_df[filter_condition].copy()

    if df_filtered.empty:
        st.warning("⚠️ 没有符合筛选条件的数据")
        return

    # 创建透视表
    df_pivot_table = create_store_pivot_table(df_filtered, selected_periods)

    if df_pivot_table.empty:
        st.warning("⚠️ 无法创建透视表，数据可能不完整")
        return

    # ==================== 全局总览指标 ====================
    st.markdown('<div class="section-title">🏁 全局总览指标</div>', unsafe_allow_html=True)

    all_period_total_gmv = df_filtered["GMV"].sum()

    if len(selected_periods) >= 2:
        latest_cycle = selected_periods[-1]
        prev_cycle = selected_periods[-2]
        current_total_gmv = df_filtered[df_filtered["统计周期"] == latest_cycle]["GMV"].sum()
        last_cycle_total_gmv = df_filtered[df_filtered["统计周期"] == prev_cycle]["GMV"].sum()
        gmv_change_amount = current_total_gmv - last_cycle_total_gmv
        gmv_change_rate = (gmv_change_amount / last_cycle_total_gmv) * 100 if last_cycle_total_gmv != 0 else 0

        # 找增长最快的店铺
        if not df_pivot_table.empty and "GMV环比%" in df_pivot_table.columns:
            growth_df = df_pivot_table[["店铺", "GMV环比%"]].copy()
            growth_df = growth_df.dropna()
            if not growth_df.empty:
                top_growth_store = growth_df.loc[growth_df["GMV环比%"].idxmax(), "店铺"]
            else:
                top_growth_store = "暂无"
        else:
            top_growth_store = "暂无"
    else:
        latest_cycle = selected_periods[0] if selected_periods else "N/A"
        current_total_gmv = all_period_total_gmv
        gmv_change_amount = 0
        gmv_change_rate = 0
        top_growth_store = "暂无"

    # 5列卡片布局
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown(f"""
        <div class="metric-card blue">
            <div class="metric-label">💰 全周期总GMV</div>
            <div class="metric-value">{format_num(all_period_total_gmv)}</div>
            <div class="metric-sub">累计销售额</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="metric-card green">
            <div class="metric-label">📈 本期({latest_cycle})GMV</div>
            <div class="metric-value">{format_num(current_total_gmv)}</div>
            <div class="metric-sub">最新周期销售额</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        delta_class = "delta-positive" if gmv_change_amount >= 0 else "delta-negative"
        delta_symbol = "▲" if gmv_change_amount >= 0 else "▼"
        st.markdown(f"""
        <div class="metric-card gray">
            <div class="metric-label">📊 GMV环比变化额</div>
            <div class="metric-value {delta_class}">{delta_symbol} {format_num(abs(gmv_change_amount))}</div>
            <div class="metric-sub">较上期变化金额</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        delta_class = "delta-positive" if gmv_change_rate >= 0 else "delta-negative"
        st.markdown(f"""
        <div class="metric-card purple">
            <div class="metric-label">📉 GMV环比变化率</div>
            <div class="metric-value {delta_class}">{gmv_change_rate:+.1f}%</div>
            <div class="metric-sub">{'增长' if gmv_change_rate >= 0 else '下降'}趋势</div>
        </div>
        """, unsafe_allow_html=True)

    with col5:
        st.markdown(f"""
        <div class="metric-card orange">
            <div class="metric-label">🏆 GMV增长TOP店铺</div>
            <div class="metric-value">{top_growth_store}</div>
            <div class="metric-sub">本期增长最快店铺</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # ==================== 图表区域 ====================
    st.markdown('<div class="section-title">📈 多周期店铺对比图表</div>', unsafe_allow_html=True)

    if not df_filtered.empty:
        col1, col2 = st.columns(2)

        with col1:
            fig = px.bar(df_filtered, x="店铺", y="GMV", color="统计周期",
                         barmode="group", title="各店铺GMV对比",
                         color_discrete_sequence=COLOR_PALETTE)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig = px.bar(df_filtered, x="店铺", y="销量", color="统计周期",
                         barmode="group", title="各店铺销量对比",
                         color_discrete_sequence=COLOR_PALETTE)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # 周期趋势图
    if selected_periods and not df_filtered.empty:
        st.markdown('<div class="section-title">📉 周期趋势分析</div>', unsafe_allow_html=True)

        cycle_summary = []
        for p in selected_periods:
            cycle_data = df_filtered[df_filtered["统计周期"] == p]
            if not cycle_data.empty:
                cycle_summary.append({
                    "周期": p,
                    "总GMV": cycle_data["GMV"].sum(),
                    "总销量": cycle_data["销量"].sum()
                })
        if cycle_summary:
            cs_df = pd.DataFrame(cycle_summary)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=cs_df["周期"], y=cs_df["总销量"], name="销量", marker_color="#70AD47"))
            fig.add_trace(go.Scatter(x=cs_df["周期"], y=cs_df["总GMV"], name="GMV", marker_color="#4472C4", yaxis="y2"))
            fig.update_layout(
                height=400,
                yaxis=dict(title="销量"),
                yaxis2=dict(title="GMV", overlaying="y", side="right")
            )
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # 单店铺趋势
    st.markdown('<div class="section-title">🏪 单店铺周度趋势</div>', unsafe_allow_html=True)

    if not df_filtered.empty:
        col1, col2 = st.columns(2)

        with col1:
            fig = px.line(df_filtered, x="统计周期", y="GMV", color="店铺",
                          markers=True, title="GMV趋势",
                          color_discrete_sequence=COLOR_PALETTE)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig = px.bar(df_filtered, x="统计周期", y="销量", color="店铺",
                         barmode="group", title="销量趋势",
                         color_discrete_sequence=COLOR_PALETTE)
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # 环比分析
    if not df_pivot_table.empty and "GMV环比%" in df_pivot_table.columns:
        st.markdown('<div class="section-title">📊 各店铺GMV环比变化率</div>', unsafe_allow_html=True)

        fig = px.bar(df_pivot_table, y="店铺", x="GMV环比%", orientation="h",
                     title="GMV环比变化率",
                     color="GMV环比%",
                     color_continuous_scale=[(0, "#d73027"), (0.3, "#f0f0f0"), (0.7, "#f0f0f0"), (1, "#1a9850")],
                     text_auto='.1f')
        fig.update_layout(height=max(400, len(df_pivot_table) * 35))
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("---")

    # 占比和排名
    if not df_pivot_table.empty:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<div class="section-title">🥧 GMV占比</div>', unsafe_allow_html=True)
            fig = px.pie(df_pivot_table, values="全周期累计GMV", names="店铺",
                         hole=0.3, title="全周期GMV占比",
                         color_discrete_sequence=COLOR_PALETTE)
            fig.update_traces(texttemplate="%{label}<br>%{percent:.1%}", textposition="inside")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            if "排名" in df_pivot_table.columns and len(selected_periods) >= 2:
                st.markdown('<div class="section-title">📈 店铺本期GMV排名</div>', unsafe_allow_html=True)
                latest_cycle = selected_periods[-1]
                if f"{latest_cycle} GMV" in df_pivot_table.columns:
                    rank_df = df_pivot_table[["排名", "店铺", f"{latest_cycle} GMV", "GMV环比%"]].copy()
                    rank_df = rank_df.sort_values("排名")
                    rank_df[f"{latest_cycle} GMV"] = rank_df[f"{latest_cycle} GMV"].apply(lambda x: f"{x:,.2f}")
                    rank_df["GMV环比%"] = rank_df["GMV环比%"].apply(lambda x: f"{x:+.1f}%")
                    rank_df.columns = ["排名", "店铺", f"{latest_cycle} GMV", "环比变化率"]
                    st.dataframe(rank_df, use_container_width=True, hide_index=True)

    st.markdown("---")

    # 店铺明细表
    if not df_pivot_table.empty:
        st.markdown('<div class="section-title">📋 店铺详细数据</div>', unsafe_allow_html=True)

        display_df = df_pivot_table.copy()
        for p in selected_periods:
            if f"{p} GMV" in display_df.columns:
                display_df[f"{p} GMV"] = display_df[f"{p} GMV"].apply(lambda x: f"{x:,.2f}")
            if f"{p} 销量" in display_df.columns:
                display_df[f"{p} 销量"] = display_df[f"{p} 销量"].apply(lambda x: f"{x:,.0f}")

        if "GMV环比变化额" in display_df.columns:
            display_df["GMV环比变化额"] = display_df["GMV环比变化额"].apply(lambda x: f"{x:+,.2f}")
        if "GMV环比%" in display_df.columns:
            display_df["GMV环比%"] = display_df["GMV环比%"].apply(lambda x: f"{x:+.2f}%")
        if "销量环比%" in display_df.columns:
            display_df["销量环比%"] = display_df["销量环比%"].apply(lambda x: f"{x:+.2f}%")
        if "全周期累计GMV" in display_df.columns:
            display_df["全周期累计GMV"] = display_df["全周期累计GMV"].apply(lambda x: f"{x:,.2f}")

        st.dataframe(display_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()