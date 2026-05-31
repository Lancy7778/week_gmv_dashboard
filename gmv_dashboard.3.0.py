import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from typing import Dict
import re
import numpy as np

# 全局周期
if 'period_order' not in st.session_state:
    st.session_state.period_order = []

COLOR_PALETTE = px.colors.qualitative.Bold

# 页面配置
st.set_page_config(
    page_title="多店铺GMV看板",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    .main .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px; padding: 20px 15px; text-align: center;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        height: 100%;
    }
    .metric-card.blue { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); }
    .metric-card.green { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
    .metric-card.orange { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
    .metric-card.purple { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
    .metric-card.gray { background: #f8f9fa; border: 1px solid #e9ecef; }
    .metric-label { color: rgba(255,255,255,0.9); font-size: 0.85rem; font-weight: 500; text-transform: uppercase; }
    .metric-value { color: white; font-size: 1.8rem; font-weight: bold; }
    .metric-sub { color: rgba(255,255,0.7); font-size: 0.75rem; }
    .delta-positive { color: #00ff88; font-weight: bold; }
    .delta-negative { color: #ff6b6b; font-weight: bold; }
    .section-title { font-size: 1.4rem; font-weight: bold; color: #1e3c72; margin: 20px 0; padding-left: 12px; border-left: 4px solid #4facfe; }
</style>
""", unsafe_allow_html=True)

def get_store_from_filename(file: Path) -> str:
    return file.stem.split("_")[0].strip()

def safe_float(value):
    if pd.isna(value): return 0.0
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value): return 0.0
        return float(value)
    try: return float(value)
    except: return 0.0

def safe_int(value):
    if pd.isna(value): return 0
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value): return 0
        return int(value)
    try: return int(float(value))
    except: return 0

def format_num(num):
    if pd.isna(num) or np.isnan(num): return "0.00"
    return f"{num:,.2f}"

def sort_period_key(period):
    nums = re.findall(r"\d+", str(period))
    if len(nums)>=4: return (int(nums[0]), int(nums[1]), int(nums[2]), int(nums[3]))
    elif len(nums)>=2: return (int(nums[0]), int(nums[1]), 0, 0)
    return (0,0,0,0)

# 加载数据：Dashboard + SKU全周期 + SKU周度
@st.cache_data(ttl=3600)
def load_all_excel(folder: str = "./各店铺周汇总报表"):
    excel_dir = Path(folder)
    excel_files = list(excel_dir.glob("*_多周汇总报表.xlsx"))
    if not excel_files:
        st.error("❌ 未找到Excel文件，请检查路径: " + folder)
        return None, None, []

    all_dashboard = []
    all_sku_total = []
    all_sku_week = []

    for f in excel_files:
        store = get_store_from_filename(f)
        try:
            # 1. Dashboard
            dash = pd.read_excel(f, sheet_name="Dashboard", engine="openpyxl")
            dash.columns = dash.columns.astype(str).str.strip()
            pcol = next((c for c in ["统计周期","周期","Period","period"] if c in dash.columns), None)
            if not pcol: continue
            for _,row in dash.iterrows():
                all_dashboard.append({
                    "店铺": store,
                    "统计周期": str(row[pcol]).strip(),
                    "GMV": safe_float(row.get("GMV",0)),
                    "销量": safe_int(row.get("销量",0))
                })

            # 2. SKU全周期汇总
            sku_t = pd.read_excel(f, sheet_name="SKU全周期汇总", engine="openpyxl")
            sku_t.columns = sku_t.columns.astype(str).str.strip()
            sku_t["店铺"] = store
            all_sku_total.append(sku_t)

            # 3. SKU周度明细
            sku_w = pd.read_excel(f, sheet_name="SKU周度明细", engine="openpyxl")
            sku_w.columns = sku_w.columns.astype(str).str.strip()
            sku_w["店铺"] = store
            all_sku_week.append(sku_w)
        except Exception as e:
            continue

    df_dash = pd.DataFrame(all_dashboard)
    df_sku_total = pd.concat(all_sku_total, ignore_index=True) if all_sku_total else pd.DataFrame()
    df_sku_week = pd.concat(all_sku_week, ignore_index=True) if all_sku_week else pd.DataFrame()

    periods = sorted(df_dash["统计周期"].unique().tolist(), key=sort_period_key)
    df_dash["统计周期"] = pd.Categorical(df_dash["统计周期"], categories=periods, ordered=True)
    return df_dash, df_sku_total, df_sku_week, periods

# 环比
def add_growth(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy().sort_values(["店铺","统计周期"])
    df["GMV"] = pd.to_numeric(df["GMV"], errors="coerce").fillna(0)
    df["销量"] = pd.to_numeric(df["销量"], errors="coerce").fillna(0)
    df["GMV环比%"] = df.groupby("店铺")["GMV"].pct_change()*100
    df["销量环比%"] = df.groupby("店铺")["销量"].pct_change()*100
    df["GMV环比%"] = df["GMV环比%"].replace([np.inf,-np.inf],0).fillna(0)
    df["销量环比%"] = df["销量环比%"].replace([np.inf,-np.inf],0).fillna(0)
    return df

# 店铺透视
def create_store_pivot(df: pd.DataFrame, periods: list):
    if len(periods)<2 or df.empty: return pd.DataFrame()
    df = df.copy()
    gmv_p = df.pivot_table(index="店铺", columns="统计周期", values="GMV", aggfunc="sum", fill_value=0).reset_index()
    sales_p = df.pivot_table(index="店铺", columns="统计周期", values="销量", aggfunc="sum", fill_value=0).reset_index()
    p_cols = [p for p in periods if p in gmv_p.columns]
    gmv_p.columns = ["店铺"] + [f"{p} GMV" for p in p_cols]
    sales_p.columns = ["店铺"] + [f"{p} 销量" for p in p_cols]
    merge = pd.merge(gmv_p, sales_p, on="店铺", how="left")
    if len(p_cols)>=2:
        prev,curr = p_cols[-2],p_cols[-1]
        merge["GMV环比变化额"] = merge[f"{curr} GMV"] - merge[f"{prev} GMV"]
        merge["GMV环比%"] = ((merge[f"{curr} GMV"] / merge[f"{prev} GMV"].replace(0,np.nan))-1)*100
        merge["GMV环比%"] = merge["GMV环比%"].fillna(0).replace([np.inf,-np.inf],0)
    merge["全周期累计GMV"] = merge[[c for c in merge.columns if "GMV" in c]].sum(axis=1)
    return merge

def main():
    st.markdown("""
    <div style='text-align:center;'><h1>📊 全店铺GMV&SKU看板</h1></div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    with st.spinner("加载数据中..."):
        df_dash, df_sku_total, df_sku_week, periods = load_all_excel()

    if df_dash is None or df_dash.empty:
        st.error("❌ 未读取到有效数据，请检查Excel文件。")
        return

    st.session_state.period_order = periods
    df_dash = add_growth(df_dash)

    with st.sidebar:
        st.markdown("### 🔍 筛选面板")
        stores = sorted(df_dash["店铺"].unique())
        selected_stores = st.multiselect("🏪 选择店铺", stores, default=stores)
        selected_periods = st.multiselect("📅 选择周期", periods, default=periods)
        top_sku_num = st.slider("🏆 Top SKU数量", 5, 50, 15)

    if not selected_stores or not selected_periods:
        st.warning("⚠️ 请选择店铺和周期")
        return

    df_filtered = df_dash[
        df_dash["店铺"].isin(selected_stores) &
        df_dash["统计周期"].isin(selected_periods)
    ]

    df_pivot = create_store_pivot(df_filtered, selected_periods)

    # ========= 全局指标 =========
    st.markdown('<div class="section-title">🏁 全局总览</div>', unsafe_allow_html=True)
    total_gmv = df_filtered["GMV"].sum()
    latest = selected_periods[-1]
    curr_gmv = df_filtered[df_filtered["统计周期"]==latest]["GMV"].sum()
    prev_gmv = df_filtered[df_filtered["统计周期"]==selected_periods[-2]]["GMV"].sum() if len(selected_periods)>=2 else 0
    delta = curr_gmv - prev_gmv
    rate = delta/prev_gmv*100 if prev_gmv else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: st.metric("全周期GMV", f"{total_gmv:,.2f}")
    with c2: st.metric("本期GMV", f"{curr_gmv:,.2f}")
    with c3: st.metric("环比变化", f"{delta:,.2f}", f"{rate:+.1f}%")
    with c4: st.metric("总销量", f"{df_filtered['销量'].sum():,}")
    with c5: st.metric("店铺数", len(selected_stores))

    st.markdown("---")

    # ========= 店铺GMV对比 =========
    st.markdown('<div class="section-title">📈 店铺GMV对比</div>', unsafe_allow_html=True)
    fig = px.bar(df_filtered, x="店铺", y="GMV", color="统计周期", barmode="group", title="各店铺GMV")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ========= Top SKU图表 =========
    st.markdown(f'<div class="section-title">🏆 Top {top_sku_num} SKU（全周期GMV）</div>', unsafe_allow_html=True)
    if not df_sku_total.empty:
        sku_filtered = df_sku_total[df_sku_total["店铺"].isin(selected_stores)].copy()
        sku_filtered["GMV"] = pd.to_numeric(sku_filtered["GMV"], errors="coerce").fillna(0)
        top_sku = sku_filtered.nlargest(top_sku_num, "GMV")
        fig = px.bar(top_sku, x="contribution sku", y="GMV", color="店铺", title=f"Top {top_sku_num} SKU GMV")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ========= SKU明细 =========
    st.markdown('<div class="section-title">📋 SKU明细</div>', unsafe_allow_html=True)
    if not df_sku_total.empty:
        display_sku = df_sku_total[df_sku_total["店铺"].isin(selected_stores)].copy()
        display_sku["GMV"] = display_sku["GMV"].apply(lambda x: f"{x:,.2f}")
        st.dataframe(display_sku, use_container_width=True)

if __name__ == "__main__":
    main()
