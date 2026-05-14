from time import strftime

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
import scipy.optimize as sco
# import matplotlib.pyplot as plt


# ----------------------
# 页面配置
# ----------------------
st.set_page_config(
    page_title="ABS投资跟进系统",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="图标.bmp"
)

# 自定义CSS样式
st.markdown("""
<style>
    /* 去掉页面顶部默认留白 */
    .block-container {
        padding-top: 0rem !important;
        margin-top: 0.3rem !important;
    }
    /* 隐藏顶部导航栏多余高度 */
    .stApp header {
        height: 0.05 !important;
        min-height: 0 !important;
        padding: 0 !important;
    }
    .main-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1e3a8a;
        margin-bottom: 0.2rem;
        text-align: center;
        margin-top: 0 !important;
    }
    .card {
        background-color: white;
        border-radius: 10px;
        padding: 1.0rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        margin-bottom: 1.5rem;
    }
    .card-title {
        font-size: 1.2rem;
        font-weight: bold;
        color: #3b82f6;
        margin-bottom: 0rem;
        display: flex;
        align-items: center;
    }
    .card-title svg {
        margin-right: 0.2rem;
    }
    .stat-value {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1e40af;
    }
    .positive {
        color: #10b981;
    }
    .negative {
        color: #ef4444;
    }
    .neutral {
        color: #6366f1;
    }
</style>
""", unsafe_allow_html=True)


# ----------------------
# 数据加载与预处理（已修复）
# ----------------------

# 1. 加载主数据 data.xlsx
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("data.xlsx")

        # 日期标准化
        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
            df = df.dropna(subset=["日期"])
            df["年月"] = df["日期"].dt.to_period("M").astype(str)

        # 本金-现金流 字段标准化（你后面要算剩余本金）
        if "本金-现金流" in df.columns:
            df["本金-现金流"] = pd.to_numeric(df["本金-现金流"], errors="coerce").fillna(0)

        return df
    except Exception as e:
        st.error(f"主数据加载失败：{str(e)}")
        st.stop()


df = load_data()


# 2. 加载交易记录
@st.cache_data
def load_trade():
    try:
        trades = pd.read_excel("交易记录.xlsx")

        # ====================== 修复核心 ======================
        # 1. 日期列名修复：你的文件是 交易日期，不是 日期
        if "交易日期" in trades.columns:
            trades.rename(columns={"交易日期": "日期"}, inplace=True)

        # 2. 日期标准化
        trades["日期"] = pd.to_datetime(trades["日期"], errors="coerce")
        trades = trades.dropna(subset=["日期"])
        trades["年月"] = trades["日期"].dt.to_period("M").astype(str)

        # 3. 交易金额清洗（必须）
        if "交易金额（元）" in trades.columns:
            trades["交易金额（元）"] = pd.to_numeric(trades["交易金额（元）"], errors="coerce")
            trades = trades[trades["交易金额（元）"] > 0]
            trades["交易金额（万元）"] = (trades["交易金额（元）"] / 10000).round(2)

        # 4. 产品1 字段预处理（你后面要画空值柱状图）
        if "产品1" in trades.columns:
            trades["产品1_为空"] = trades["产品1"].isnull()

        # # 5. 资产类型字段标准化
        # if "资产类型二" in trades.columns:
        #     trades.rename(columns={"资产类型二": "资产类型"}, inplace=True)

        return trades
    except Exception as e:
        st.error(f"交易记录加载失败：{str(e)}")
        st.stop()


trades = load_trade()

# 3. 加载项目库
@st.cache_data
def load_projects():
    try:
        projects = pd.read_excel("项目库.xlsx")
        return projects
    except Exception as e:
        st.error(f"项目库加载失败：{str(e)}")
        st.stop()

df_projects = load_projects()


# =========== 关联data和项目库==============
# 标准化data表的证券名称
# df["证券名称"] = df["证券名称"].str.strip().upper()
# 通过“证券名称”左关联（保留data所有数据，关联项目库的补充信息）
df_merged = pd.merge(
    df,  # 主表：data
    df_projects,  # 关联表：项目库
    on="证券名称",  # 关联字段：证券名称
    how="left",  # 左连接：保留data所有数据，项目库无匹配则为空
    suffixes=("", "_项目库")  # 重复字段加后缀区分
)


# 1. 注入 iconfont 样式（必须加）
st.markdown("""
<style>
/* iconfont 图标样式 */
.iconfont {
    font-family: "iconfont" !important;
    font-size: 18px;
    font-style: normal;
    margin-right: 8px;
    color: #4E5969;
}
/* 侧边栏菜单文字样式 */
.sidebar-menu-item {
    font-size: 15px;
    padding: 8px 0;
    color: #1F2937;
}
</style>
""", unsafe_allow_html=True)

# ----------------------
# 侧边栏配置
# ----------------------
with st.sidebar:
    st.title("投资分析菜单")

    # 👇 直接平展展示菜单，无折叠、无缩进，纯 side menu 样式
    menu = st.radio(
        label="",  # 隐藏标题
        options=[
            "核心投资指标",
            "产品分析",
            "资产分析",
            "交易记录",
        ],
        label_visibility="collapsed"  # 完全隐藏 label，界面整洁
    )

    st.markdown("---")
    # st.title("筛选条件")


    df_filtered = df_merged.copy()
    trades_filtered = trades.copy()

    # # 时间范围筛选
    # if "日期" in df_filtered.columns:
    #     min_date = df_filtered["日期"].min()
    #     max_date = df_filtered["日期"].max()
    #     selected_dates = st.date_input(
    #         "时间范围筛选",
    #         value=[min_date, max_date],
    #         min_value=min_date,
    #         max_value=max_date
    #     )
    #     # 应用时间筛选
    #     mask = (df_filtered["日期"] >= pd.to_datetime(selected_dates[0])) & \
    #            (df_filtered["日期"] <= pd.to_datetime(selected_dates[1]))
    #     df_filtered = df_filtered[mask].copy()
    #     trades_filtered = trades.loc[mask].copy()
    # else:
    #     df_filtered = df_filtered.copy()
    #     trades_filtered = trades.copy()
    #
    # # 资产大类筛选
    # if "资产类型二" in df_filtered.columns:
    #     asset_types = df_filtered["资产类型二"].dropna().unique()
    #     selected_assets = st.multiselect(
    #         "资产大类筛选",
    #         options=asset_types,
    #         default=asset_types
    #     )
    #     df_filtered = df_filtered[df_filtered["资产类型"].isin(selected_assets)]
    #     trades_filtered = trades_filtered[trades_filtered["资产类型"].isin(selected_assets)]
    # else:
    #     df_filtered = df_filtered.copy()
    #     trades_filtered = trades.copy()
    #
    # # 资产类型筛选
    # if "资产类型" in df_filtered.columns:
    #     asset_types_2 = df_projects["资产类型"].dropna().unique()
    #     selected_assets = st.multiselect(
    #         "资产类型筛选",
    #         options=asset_types_2,
    #         default=asset_types_2
    #     )
    #     df_filtered = df_filtered[df_filtered["资产类型二"].isin(selected_assets)]
    #     trades_filtered = trades_filtered[trades_filtered["资产类型"].isin(selected_assets)]
    #
    # else:
    #     df_filtered = df_filtered.copy()
    #     trades_filtered = trades.copy()
    #
    # # 产品筛选（优先用关联表的产品名称）
    # if "产品名称" in df_filtered.columns:
    #     products = df_filtered["产品名称"].dropna().unique()
    #     selected_product = st.selectbox(
    #         "产品筛选（可选）",
    #         options=["全部产品"] + list(products),
    #         index=0
    #     )
    #     if selected_product != "全部产品":
    #         df_filtered = df_filtered[df_filtered["产品名称"] == selected_product]
    #         trades_filtered = trades_filtered[trades_filtered["产品2"] == selected_product]
    #
    # # 项目库维度筛选
    # if "项目名称" in df_projects.columns:
    #     projects_list = df_projects["项目名称"].dropna().unique()
    #     selected_project = st.selectbox(
    #         "项目库维度筛选（可选）",
    #         options=["全部项目"] + list(projects_list),
    #         index=0
    #     )
    #     if selected_project != "全部项目":
    #         df_filtered = df_filtered[df_filtered["项目名称"] == selected_project]
    #         trades_filtered = trades_filtered[trades_filtered["项目名称"] == selected_project]

    # st.info(f"当前数据：{len(df_filtered)} 条 | 交易记录：{len(trades_filtered)} 条")

# ----------------------
# 核心统计计算
# ----------------------
def calculate_statistics(df):
    stats = {}

    # 基础统计
    stats["total_records"] = len(df)
    # 日期用data表的日期
    date_col = "日期" if "日期" in df.columns else "日期"
    if date_col in df.columns and not df[date_col].isnull().all():
        stats["date_range"] = f"{df[date_col].min().strftime('%Y-%m-%d')} 至 {df[date_col].max().strftime('%Y-%m-%d')}"
    else:
        stats["date_range"] = "无时间数据"

    # 交易金额统计（优先用关联表的交易金额）
    trade_col = "交易金额（元）" if "交易金额（元）" in df.columns else "交易金额（元）"
    if trade_col in df.columns:
        trade_amount = df[trade_col].dropna()
        stats["trade"] = {
            "total": trade_amount.sum(),
            "avg": trade_amount.mean(),
            "max": trade_amount.max(),
            "min": trade_amount.min(),
            "count": len(trade_amount)
        }

    # 投资金额统计（本金-现金流）
    invest_col = "本金-现金流" if "本金-现金流" in df.columns else "本金-现金流"
    if invest_col in df.columns:
        invest_amount = df[invest_col].dropna()
        stats["invest"] = {
            "total": invest_amount.sum(),
            "avg": invest_amount.mean(),
            "max": invest_amount.max(),
            "min": invest_amount.min(),
            "count": len(invest_amount)
        }


    # 投资现金流统计
    cash_col = "投资现金流" if "投资现金流" in df.columns else "投资现金流"
    if cash_col in df.columns:
        cash_flow = df[cash_col].dropna()
        positive_cash = cash_flow[cash_flow > 0]
        negative_cash = cash_flow[cash_flow < 0]
        stats["cash_flow"] = {
            "total": cash_flow.sum(),
            "avg": cash_flow.mean(),
            "max_inflow": positive_cash.max() if len(positive_cash) > 0 else 0,
            "max_outflow": negative_cash.min() if len(negative_cash) > 0 else 0,
            "total_inflow": positive_cash.sum(),
            "total_outflow": negative_cash.sum(),
            "inflow_count": len(positive_cash),
            "outflow_count": len(negative_cash)
        }

    # 资产大类统计（适配关联表的资产类型）
    asset_col = "资产类型二" if "资产类型二" in df.columns else "资产类型二"
    if asset_col in df.columns and invest_col in df.columns:
        asset_type_principal = df.groupby(asset_col)[invest_col].sum()
        asset_type_principal_abs = (asset_type_principal.abs() / 10000).round(2)
        stats["asset_type_principal_dist"] = asset_type_principal_abs.to_dict()

    # # 资产类型统计（适配关联表的资产类型）
    # asset_small_type = "资产类型" if "资产类型" in df.columns else "资产类型"
    # if asset_small_type in df.columns and invest_col in df.columns:
    #     small_type_prin = df.groupby(asset_small_type)[invest_col].sum()
    #     small_type_prin_abs = (small_type_prin.abs() / 10000).round(2)
    #     stats["small_type_dist"] = small_type_prin_abs.to_dict()

    # 产品统计
    product_col = "产品名称" if "产品名称" in df.columns else "产品名称"
    if product_col in df.columns:
        product_stats = {
            "total_products": df[product_col].nunique(),
            "top_products": df[product_col].value_counts().head(10).to_dict()
        }
        stats["product"] = product_stats

    # 证券统计（优先用关联后的证券名称）
    security_col = "证券名称" if "证券名称" in df.columns else "证券名称"
    if security_col in df.columns:
        security_stats = {
            "total_securities": df[security_col].nunique(),
            "top_securities": df[security_col].value_counts().head(10).to_dict()
        }
        stats["security"] = security_stats

    # 时间分布统计
    month_col = "年月" if "年月" in df.columns else "年月"
    if month_col in df.columns:
        monthly_dist = df[month_col].value_counts().sort_index()
        stats["monthly_dist"] = monthly_dist.to_dict()


    return stats

stats = calculate_statistics(df_filtered)

# 核心交易记录计算
def calculate_trade_records(df):
    trade_records = {}
    trade_records["total_records"] = len(df)
    trade_records["total_amount"] = df["交易金额（元）"].sum() if "交易金额（元）" in df.columns else 0
    trade_records["total_count"] = df["交易金额（元）"].count() if "交易金额（元）" in df.columns else 0
    trade_records["total_average"] = trade_records["total_amount"] / trade_records["total_count"] if trade_records["total_count"] > 0 else 0
    trade_records["total_max"] = df["交易金额（元）"].max() if "交易金额（元）" in df.columns else 0
    trade_records["total_min"] = df["交易金额（元）"].min() if "交易金额（元）" in df.columns else 0
    trade_records["total_average_per_day"] = trade_records["total_amount"] / len(df["日期"].unique()) if len(df["日期"].unique()) > 0 else 0
    trade_records["total_average_per_month"] = trade_records["total_amount"] / len(df["年月"].unique()) if len(df["年月"].unique()) > 0 else 0
    trade_records["product1_empty"] = len(df[df["产品1"].isnull()]) if "产品1" in df.columns else 0
    return trade_records

trade_records = calculate_trade_records(trades_filtered)

# ----------------------
# 主页面内容
# ----------------------
st.markdown("<h1 class='main-header'>ABS投资跟进</h1>", unsafe_allow_html=True)
st.markdown(f"""
<div style='text-align: right; font-size: 14px; margin-bottom: 10px; font-weight: normal;'>
    数据时间范围：{stats['date_range']}
</div>
""", unsafe_allow_html=True)
# 超紧凑分割线（间距极小）
st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# 1. 核心指标卡片（4列布局）
if menu == "核心投资指标":

    # st.subheader("核心投资指标")
    col1, col2, col3, col4 = st.columns(4,gap = "small")

    # 1. 存续投资金额（万元）
    with col1:

        invest_total = stats["invest"]["total"] if "invest" in stats else 0
        value_invest_total = f"{-invest_total / 10000:,.2f}"
        st.markdown(f"""
        <div class='card'>
            <div class='card-title' style='color: black;'>存续投资金额（万元）</div>
            <div style='font-size: 28px; font-weight: normal; color: black; margin-top: 5px;'>
                {value_invest_total}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 2. 累计投资规模卡片
    with col2:
        # cashflow中的流出求和
        max_inflow = stats["cash_flow"]["max_inflow"] if "cash_flow" in stats else 0
        max_outflow = stats["cash_flow"]["max_outflow"] if "cash_flow" in stats else 0
        total_outflow = stats["cash_flow"]["total_outflow"] if "cash_flow" in stats else 0
        value_total_outflow = f"{-total_outflow / 10000:,.2f}"

        st.markdown(f"""
        <div class='card'>
            <div class='card-title' style='color: black;'>累计投资金额（万元）</div>
            <div style='font-size: 28px; font-weight: normal; color: black; margin-top: 5px;'>
                {value_total_outflow}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 3. 产品数量卡片
    with col3:
        total_products = stats["product"]["total_products"] if "product" in stats else 0
        value_total_products = f"{total_products:,}"
        st.markdown(f"""
        <div class='card'>
            <div class='card-title' style='color: black;'>产品数量</div>
            <div style='font-size: 28px; font-weight: normal; color: black; margin-top: 5px;'>
                {value_total_products:}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 4. 证券信息卡片
    with col4:
        total_securities = stats["security"]["total_securities"] if "security" in stats else 0
        value_total_securities = f"{total_securities:,}"
        st.markdown(f"""
        <div class='card'>
            <div class='card-title' style='color: black;'>证券数量</div>
            <div style='font-size: 28px; font-weight: normal; color: black; margin-top: 5px;'>
                {value_total_securities:}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 超紧凑分割线（间距极小）
    st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

    # 2. 图表区域（2列布局）
    # st.subheader("资产分布")
    chart_col1, chart_col2 = st.columns(2)

    # 2.1 资产类型分布饼图

    with chart_col1:
        st.markdown("""
        <div class='card'>
            <div class='card-title' style='color: black;'>
                资产类型分布
            </div>
        """, unsafe_allow_html=True)

        if "asset_type_principal_dist" in stats and len(stats["asset_type_principal_dist"]) > 0:
            asset_data = pd.DataFrame({
                "资产类型": list(stats["asset_type_principal_dist"].keys()),
                "剩余本金（万元）": list(stats["asset_type_principal_dist"].values()),
            })

            # 计算占比
            asset_data["占比(%)"] = (asset_data["剩余本金（万元）"] / asset_data["剩余本金（万元）"].sum() * 100).round(1)

            fig1 = px.pie(
                asset_data,
                values="剩余本金（万元）",
                names="资产类型",
                hole=0.5,
                color_discrete_sequence=["#3b82f6", "#ef4444", "#10b981", "#f59e0b"]
            )
            fig1.update_traces(textposition='inside', textinfo='percent+label')
            fig1.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=300)
            st.plotly_chart(fig1, width="stretch")

            # 表格格式化：保留千分号和两位小数
            asset_data["剩余本金（万元）"] = asset_data["剩余本金（万元）"].apply(lambda x: f"{x:,.2f}")
            st.dataframe(
                asset_data[["资产类型", "剩余本金（万元）", "占比(%)"]],
                width="stretch",
                hide_index=True
            )
        else:
            st.warning("暂无资产类型数据")

        st.markdown("</div>", unsafe_allow_html=True)


    # ====================== 资产统计 ======================
    with chart_col2:



   # 2.2 月度交易趋势图
        # 把卡片开头 + 标题 + 图标 写在一起，就正常了
        st.markdown("""
        <div class='card'>
            <div class='card-title' style='color: black;'>
                月度交易趋势
            </div>
        """, unsafe_allow_html=True)

        # 筛选产品1为空的记录并处理数据
        trades_null_product1 = trades_filtered[trades_filtered['产品1'].isnull()].copy()

        if len(trades_null_product1) > 0:
            # 按年月分组统计
            trades_null_product1['交易年月'] = trades_null_product1['日期'].dt.to_period('M').astype(str)
            monthly_stats = trades_null_product1.groupby('交易年月').agg({
                '交易金额（元）': ['sum', 'count']
            }).round(2)
            monthly_stats.columns = ['交易金额总和（元）', '交易笔数']
            monthly_stats = monthly_stats.sort_index().reset_index()

            # 转换为万元单位
            monthly_stats['交易金额总和（万元）'] = (monthly_stats['交易金额总和（元）'] / 10000).round(2)

            # 1. 月度交易金额柱状图（默认近一年 + 纵轴自适应）
            fig_amount = px.bar(
                monthly_stats,
                x='交易年月',
                y='交易金额总和（万元）',
                color_discrete_sequence=["#3b82f6"],
            )
            fig_amount.update_traces(
                textposition='outside',
                texttemplate='%{y:.1f}',
                hovertemplate='月份: %{x}<br>交易金额: %{y:.2f} 万元'
            )
            fig_amount.update_layout(
                margin=dict(l=0, r=0, t=20, b=80),
                height=300,
                xaxis_title="交易月份",
                yaxis_title="交易金额（万元）",
                # 👇 纵轴完全自适应（自动匹配数据最大值）
                yaxis=dict(
                    automargin=True,
                    fixedrange=False
                )
            )

            # 👇 核心：默认只显示【近12个月】
            last_date = pd.to_datetime(monthly_stats['交易年月'].max())
            one_year_ago = last_date - pd.DateOffset(years=1)

            # 筛选近一年数据
            mask = pd.to_datetime(monthly_stats['交易年月']) >= one_year_ago
            display_months = monthly_stats[mask]['交易年月'].tolist()

            fig_amount.update_xaxes(
                tickangle=-45,
                categoryorder='array',
                categoryarray=monthly_stats['交易年月'],
                range=[display_months[0], display_months[-1]] if len(display_months) > 0 else None
            )

            st.plotly_chart(fig_amount, width="stretch")


        else:
            st.warning("暂无产品1为空的交易记录")

        st.markdown("</div>", unsafe_allow_html=True)

    # 以时间为维度，以年为单位，计算截止不同时间点的投资本金余额
    # 确保字段存在
    if all(col in df_filtered.columns for col in ["日期", "资产类型二", "本金-现金流"]):
        df_time = df_filtered.copy()

        # 1. 提取年份（横坐标按年）
        df_time["年份"] = df_time["日期"].dt.year

        # 2. 按【年份】求和本金现金流
        df_year = df_time.groupby(["年份", "资产类型二"])["本金-现金流"].sum().reset_index()

        # 3. 计算累计投资本金余额
        df_year = df_year.sort_values(["资产类型二", "年份"])
        df_year["投资本金余额（万元）"] = df_year.groupby("资产类型二")["本金-现金流"].cumsum().abs() / 10000

        if df_year.empty:
            st.warning("暂无大于0的投资本金余额数据")
        else:
            import plotly.express as px

            # 超紧凑分割线（间距极小）
            st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)
            st.markdown("""
            <div class='card'>
                <div class='card-title' style='color: black;'>
                    规模余额
                </div>
            """, unsafe_allow_html=True)
            fig = px.bar(
                df_year,
                x="年份",  # 按年
                y="投资本金余额（万元）",
                color="资产类型二",
                barmode="stack",
                text_auto=",.2f"

            )
            # 强制X轴显示为整数年份
            fig.update_xaxes(
                tickmode="array",
                tickvals=sorted(df_year["年份"].unique()),
                tickformat="d"
            )
            fig.update_layout(
                height=450,
                xaxis_title="年份",
                yaxis_title="投资本金余额（万元）",
                margin=dict(l=0, r=0, t=40, b=0)
            )
            st.plotly_chart(fig, width="stretch")

    else:
        st.warning("缺少 日期/产品名称/本金-现金流 字段，无法生成图表")

    st.markdown("</div>", unsafe_allow_html=True)

    # 按产品分组，计算 本金-现金流 总和（投资本金余额）
    if "产品名称" in df_filtered.columns and "本金-现金流" in df_filtered.columns and "日期" in df_filtered.columns:
        product_balance = df_filtered.groupby("产品名称").agg(
            投资本金余额=("本金-现金流", "sum"),  # 求和
            最早投资日期=("日期", "min")  # 取最早日期（正确！）
        ).reset_index()
        product_balance.columns = ["产品名称", "投资本金余额", "最早投资日期"]

        # 👇 只保留 投资本金余额 > 0 的产品
        product_balance = product_balance[product_balance["投资本金余额"] < 0]

        # 转为正数（更符合规模展示）
        product_balance["投资本金余额（万元）"] = product_balance["投资本金余额"].round(2).abs() / 10000

        # 计算占比
        total_balance = product_balance["投资本金余额（万元）"].sum()
        product_balance["占比(%)"] = (product_balance["投资本金余额（万元）"] / total_balance * 100).round(2)

        # 计算各产品的最小日期
        # product_balance["最早投资日期"] = product_balance["日期"].min()

        # 按金额从大到小排序
        product_balance = product_balance.sort_values("投资本金余额（万元）", ascending=False)

        # 产品规模柱状图，x轴只显示前20个
        # 超紧凑分割线（间距极小）
        st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title' style='color: black;'>
                产品存续规模Top20
            </div>
        """, unsafe_allow_html=True)
        product_balance = product_balance.head(20)
        fig = px.bar(
            product_balance,
            x="产品名称",
            y="投资本金余额（万元）",
            color="投资本金余额（万元）",
            text_auto=",.2f",
            color_continuous_scale="Blues"

        )
        fig.update_layout(
            margin=dict(l=0, r=0, t=40, b=0),
            height=450,
            xaxis_title="",
            yaxis_title="投资本金余额（万元）",
            showlegend=False
        )
        fig.update_xaxes(tickangle=-45)
        fig.update_coloraxes(showscale=False)
        st.plotly_chart(fig, width="stretch")

        # # 产品规模详情表格
        # st.markdown("### 产品投资规模详情")
        # # 显示：产品名称、投资本金余额（万元）、最早持仓时间
        # product_balance = product_balance[["产品名称", "投资本金余额（万元）", "最早投资日期"]]
        # st.dataframe(
        #     product_balance.style.format({
        #         "投资本金余额（万元）": "{:,.2f}",
        #         "最早投资日期": lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) else ""
        #     }),
        #     width="stretch",
        #     hide_index=True
        # )

    else:
        st.warning("暂无产品规模数据")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("---")

    # ----------------------
    # 3. 产品分析（按本金-现金流求和）
    # ----------------------
elif menu == "产品分析":
    # st.subheader("产品分析")



    # ----------------------
    # 📋 各产品最新持仓收益表（带产品+证券筛选）
    # ----------------------
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>
            产品持仓及收益明细
        </div>
    """, unsafe_allow_html=True)

    # 检查字段
    must_fields = [
        "产品名称", "证券名称", "日期",
        "持仓份额（万份）",
        "利息-现金流", "本金-现金流","投资现金流"
    ]
    missing_fields = [f for f in must_fields if f not in df_filtered.columns]

    if missing_fields:
        st.error(f"缺少关键字段：{', '.join(missing_fields)}")
    else:
        # 以持仓份额降序

        df_copy = df_filtered.copy().sort_values("本金-现金流", ascending=True)

        # ============================
        # ✅ 表格筛选器：产品 + 证券，模糊搜索
        # ============================
        col1, col2 = st.columns(2)
        with col1:
            product_list = st.text_input("产品名称", placeholder="输入产品关键词...")
            # selected_product = st.selectbox("筛选产品", product_list)
        with col2:
            security_list = st.text_input("证券名称", placeholder="输入证券关键词...")
            # selected_security = st.selectbox("筛选证券", security_list)

        # 应用筛选
        if product_list:
            df_copy = df_copy[df_copy["产品名称"].str.contains(product_list, na=False, case=False)]
        if security_list:
            df_copy = df_copy[df_copy["证券名称"].str.contains(security_list, na=False, case=False)]

        if df_copy.empty:
            st.warning("⚠️ 当前筛选条件下无数据")
        else:
            # ======================================================================
            # 步骤1：取【产品+证券】最新时间点数据
            # ======================================================================
            df_latest_date = df_copy.groupby(["产品名称", "证券名称"])["日期"].max().reset_index()
            df_latest = df_copy.merge(df_latest_date, on=["产品名称", "证券名称", "日期"], how="inner")

            # ======================================================================
            # 步骤2：✅ 持仓份额 = 【产品+证券】的 交易份额 求和
            # ======================================================================
            share_sum = df_copy.groupby(["产品名称", "证券名称"])["交易份额（万份）"].sum().reset_index()
            share_sum.rename(columns={"交易份额（万份）": "持仓份额"}, inplace=True)
            df_latest = df_latest.merge(share_sum, on=["产品名称", "证券名称"], how="left")

            # ======================================================================
            # 步骤3：投资本金余额 = 最新时间点投资本金余额
            # ======================================================================
            balance_sum = df_copy.groupby(["产品名称", "证券名称"])["本金-现金流"].sum().reset_index()
            balance_sum.rename(columns={"本金-现金流": "投资本金余额"}, inplace=True)
            df_latest = df_latest.merge(balance_sum, on=["产品名称", "证券名称"], how="left")

            # ======================================================================
            # 步骤4：收益金额 = 利息-现金流 总和
            # ======================================================================
            interest_total = df_copy.groupby(["产品名称", "证券名称"])["利息-现金流"].sum().reset_index()
            interest_total.rename(columns={"利息-现金流": "收益金额"}, inplace=True)
            df_latest = df_latest.merge(interest_total, on=["产品名称", "证券名称"], how="left")

            # ======================================================================
            # 步骤5：XIRR 计算
            # ======================================================================
            try:

                def xirr(cashflows, dates, guess=0.1):
                    try:
                        import scipy.optimize as sco
                        def npv(rate):
                            return sum(
                                [cf / (1 + rate) ** ((d - dates[0]).days / 365.0) for cf, d in zip(cashflows, dates)])

                        return sco.newton(npv, guess)
                    except:
                        return None

                def calc_xirr(group):
                    values = group["投资现金流"].fillna(0).tolist()
                    dates = group["日期"].tolist()
                    if len(values) < 2:
                        return None
                    return xirr(values, dates)

                xirr_result = df_copy.groupby(["产品名称", "证券名称"]).apply(calc_xirr).reset_index()
                xirr_result.columns = ["产品名称", "证券名称", "XIRR"]
                df_latest = df_latest.merge(xirr_result, on=["产品名称", "证券名称"], how="left")
            except:
                df_latest["XIRR"] = None
            # 按产品名称+证券名称分组聚合
            df_latest = df_latest.groupby(["产品名称", "证券名称"], as_index=False).agg({
                "持仓份额": "sum",
                "投资本金余额": "sum",
                "收益金额": "sum",
                "XIRR": "first"
            })

            # ======================================================================
            # 最终表格
            # ======================================================================
            final_table = df_latest[[
                "产品名称", "证券名称", "持仓份额", "投资本金余额", "收益金额", "XIRR"
            ]].copy()

            # 格式化
            final_table["持仓份额"] = final_table["持仓份额"].astype(float).round(2)
            final_table["投资本金余额"] = final_table["投资本金余额"].astype(float).round(2).abs()
            final_table["收益金额"] = final_table["收益金额"].astype(float).round(2)
            final_table["XIRR"] = final_table["XIRR"].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "—")

            # 显示表格
            # 只显示持仓份额 > 0的，并按照投资本金余额排序
            final_table = final_table[final_table["持仓份额"] > 0].sort_values(by="投资本金余额", ascending=False)
            st.dataframe(
                final_table.style.format({
                    "持仓份额": "{:,.2f}",
                    "投资本金余额": "{:,.2f}",
                    "收益金额": "{:,.2f}",
                }),
                width="stretch", height=450
            )
            # ---------------------
            # ✅ 汇总求和（交易份额 + 交易金额）
            # ---------------------
            total_share_final_table = final_table["持仓份额"].sum()
            total_amount_final_table = final_table["投资本金余额"].sum()
            total_profit_final_table = final_table["收益金额"].sum()


            # ---------------------
            # ✅ 底部汇总行（卡片内）
            # ---------------------
            st.markdown(f"""
            <div style="margin-top:10px; font-size:15px; font-weight:bold; color:#1e40af;">
                合计： 持仓份额 = {total_share_final_table:,.2f} 万份  |  投资本金余额 = {total_amount_final_table:,.2f} 万元  |  收益金额 = {total_profit_final_table:,.2f} 万元
            </div>
            """, unsafe_allow_html=True)


    st.markdown("</div>", unsafe_allow_html=True)


# 数据来自“项目库.xlsx"
# 展示项目清单
# ----------------------
# 资产分析页面（含项目清单）
# ----------------------

# ----------------------
# 资产分析页面（含发行规模转换+固定日期格式）
# ----------------------
elif menu == "资产分析":
    # st.markdown("<h2 class='main-header' style='text-align:left; margin-bottom:1rem;'>资产分析</h2>",
    #             unsafe_allow_html=True)







    # ====================== 带筛选的项目清单模块（优化版） ======================
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>资产清单</div>
    """, unsafe_allow_html=True)

    # 1. 加载项目库Excel
    try:
        # df_projects = pd.read_excel("项目库.xlsx", sheet_name=0)  # 按实际Sheet名调整
        df_projects_copy = df_projects.copy()

        # 2. 多字段模糊筛选器（保持原有布局）
        st.markdown("<div style='margin-bottom:15px;'>", unsafe_allow_html=True)
        # 第一行：项目名称 + 证券名称
        filter_col1, filter_col2 = st.columns(2, gap="small")
        with filter_col1:
            search_project = st.text_input("项目名称", placeholder="输入项目关键词...")
        with filter_col2:
            search_security = st.text_input("证券名称", placeholder="输入证券关键词...")

        # 第二行：资产类型 + 主体所属 + 资产类型二
        filter_col3, filter_col4, filter_col5 = st.columns(3, gap="small")
        with filter_col3:
            search_asset1 = st.text_input("资产类型", placeholder="输入资产类型关键词...")
        with filter_col4:
            search_subject = st.text_input("主体所属", placeholder="输入主体关键词...")
        with filter_col5:
            search_asset2 = st.text_input("资产类型二", placeholder="输入资产类型二关键词...")
        st.markdown("</div>", unsafe_allow_html=True)

        # 3. 应用模糊筛选
        if search_project:
            df_projects_copy = df_projects_copy[
                df_projects_copy["项目名称"].str.contains(search_project, na=False, case=False)
            ]
        if search_security:
            df_projects_copy = df_projects_copy[
                df_projects_copy["证券名称"].str.contains(search_security, na=False, case=False)
            ]
        if search_asset1:
            df_projects_copy = df_projects_copy[
                df_projects_copy["资产类型"].str.contains(search_asset1, na=False, case=False)
            ]
        if search_subject:
            df_projects_copy = df_projects_copy[
                df_projects_copy["主体所属"].str.contains(search_subject, na=False, case=False)
            ]
        if search_asset2:
            df_projects_copy = df_projects_copy[
                df_projects_copy["资产类型二"].str.contains(search_asset2, na=False, case=False)
            ]

        # 4. 核心优化：数据格式化（发行规模单位转换+固定日期格式）
        if not df_projects_copy.empty:
            # ---------------------- 优化1：发行规模（元→万元）转换 ----------------------
            # if "发行规模（元）" in df_projects_copy.columns:
            #     # 1. 单位转换：元 ÷ 10000 = 万元，保留2位小数
            #     df_projects_copy["发行规模（万元）"] = df_projects_copy["发行规模（元）"].apply(
            #         lambda x: round(x / 10000, 2) if pd.notna(x) and isinstance(x, (int, float)) else "—"
            #     )
            #     # 先确保是数值类型，非数值转为 NaN
            #     df_projects_copy["发行规模（万元）"] = pd.to_numeric(df_projects_copy["发行规模（万元）"], errors="coerce")
            #
            #     # 格式化：保留2位小数 + 千分位分隔符，空值显示“—”
            #     df_projects_copy["发行规模（万元）"] = df_projects_copy["发行规模（万元）"].apply(
            #         lambda x: f"{x:,.2f}" if pd.notna(x) else "—"
            #     )
                # 2. 删除原“发行规模（元）”列，只保留万元列
            df_projects_copy = df_projects_copy.drop(columns=["发行规模（元）"])

            # ---------------------- 优化2：日期格式强制为 yyyy-mm-dd ----------------------
            # 识别所有含“日期”的列（如发行日期、到期日期等）
            date_cols = [col for col in df_projects_copy.columns if "预期到期日" in col or "清算日期" in col]
            for col in date_cols:
                # 强制转换为datetime，无效日期显示空值，最终格式固定为yyyy-mm-dd
                df_projects_copy[col] = pd.to_datetime(
                    df_projects_copy[col],
                    errors="coerce"  # 无法转换的日期设为NaT
                ).dt.strftime("%Y-%m-%d")  # 固定格式输出
                # 空值（NaT）显示为“—”，避免显示“NaT”
                df_projects_copy[col] = df_projects_copy[col].replace("NaT", "—")

            # ---------------------- 原有：预期收益率百分比格式化 ----------------------
            if "预期收益率" in df_projects_copy.columns:
                df_projects_copy["预期收益率"] = df_projects_copy["预期收益率"].apply(
                    lambda x: f"{x * 100:.2f}%" if pd.notna(x) and isinstance(x, (int, float)) else "—"
                )

            # 5. 显示美化表格
            st.dataframe(
                df_projects_copy.style
                .set_properties(**{"text-align": "left", "font-size": "13px", "white-space": "nowrap"})
                .set_table_styles([{"selector": "th", "props": [("text-align", "left"), ("font-weight", "bold")]}]),
                width="stretch",
                height=350,
                hide_index=True
            )

            # 6. 底部统计（新增发行规模汇总）
            total_filtered = len(df_projects_copy)
            stats_text = f"筛选后项目总数：<span style='font-weight:bold; color:#1e40af;'>{total_filtered} 个</span>"

            # 新增：发行规模（万元）汇总（仅统计有效数值）
            if "发行规模（万元）" in df_projects_copy.columns:
                # 提取有效数值（排除“—”）
                valid_scale = df_projects_copy["发行规模（万元）"][df_projects_copy["发行规模（万元）"] != "—"]
                if not valid_scale.empty:
                    total_scale = valid_scale.astype(float).sum()
                    stats_text += f" | 总发行规模：<span style='font-weight:bold; color:#1e40af;'>{total_scale:,.2f} 万元</span>"

            # 预期收益率统计（保留原有）
            if "预期收益率" in df_projects.columns:
                valid_yields = df_projects[
                    df_projects["预期收益率"].apply(lambda x: pd.notna(x) and isinstance(x, (int, float)))]
                if not valid_yields.empty:
                    avg_yield = valid_yields["预期收益率"].mean()
                    stats_text += f" | 平均预期收益率：<span style='font-weight:bold; color:#1e40af;'>{avg_yield * 100:.2f}%</span>"

            st.markdown(f"""
            <div style="margin-top:10px; font-size:14px; color:#333;">
                {stats_text}
            </div>
            """, unsafe_allow_html=True)

        else:
            st.warning("⚠️ 当前筛选条件下无匹配项目，请调整关键词后重试")

    except FileNotFoundError:
        st.error("❌ 未找到项目库.xlsx文件，请确认文件在项目根目录（与app.py同级）")
    except Exception as e:
        st.error(f"❌ 项目清单加载失败：{str(e)}")
        st.info("请检查Excel文件格式（如Sheet名称、“发行规模（元）”字段是否存在）")

    st.markdown("</div>", unsafe_allow_html=True)  # 项目清单卡片结束
    st.markdown("<br>", unsafe_allow_html=True)

    # ====================== 原有资产分析模块（保持不变） ======================
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>资产分布分析</div>
    """, unsafe_allow_html=True)

    # （此处插入你原有的资产分布图表、数据展示代码）
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------
# 交易记录页面（带筛选 + 正确字段 + 美观展示）
# ---------------------
elif menu == "交易记录":
    # st.markdown("<h2 class='main-header' style='text-align:left; margin-bottom:1rem;'>交易记录</h2>", unsafe_allow_html=True)

    # 筛选区域
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>
            交易记录
        </div>
    """, unsafe_allow_html=True)

    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)

    with filter_col1:
        min_date = trades_filtered["日期"].min()
        max_date = trades_filtered["日期"].max()
        date_range = st.date_input("交易时间", value=[min_date, max_date])


    with filter_col2:
        securities = st.text_input("证券名称", placeholder="模糊搜索...")
        # selected_sec = st.selectbox("证券名称", securities)

    with filter_col3:
        sellers = st.text_input("卖出方", placeholder="模糊搜索...")
        # selected_seller = st.selectbox("卖出主体", sellers)


    with filter_col4:
        buyers = st.text_input("买入方", placeholder="模糊搜索...")
        # selected_buyer = st.selectbox("买入主体", buyers)

    st.markdown("</div>", unsafe_allow_html=True)

    # 应用筛选
    trades_show = trades_filtered.copy()

    # 时间筛选
    rades_show = trades_show[
        (trades_show["日期"] >= pd.to_datetime(date_range[0])) &
        (trades_show["日期"] <= pd.to_datetime(date_range[1]))
    ]

    # 证券筛选
    if securities:
        trades_show = trades_show[trades_show["证券名称"].str.contains(securities, na=False, case=False)]

    # 买入方筛选
    if buyers:
        trades_show = trades_show[trades_show["买入主体"].str.contains(buyers, na=False, case=False)]

    # 卖出方筛选
    if sellers:
        trades_show = trades_show[trades_show["卖出主体"].str.contains(sellers, na=False, case=False)]

    # # 展示交易记录表
    # st.markdown("""
    # <div class='card'>
    #     <div class='card-title' style='color: black;'>交易记录明细</div>
    # """, unsafe_allow_html=True)

    # 格式化日期为yyyy-mm-dd
    trades_show["日期"] = trades_show["日期"].dt.strftime("%Y-%m-%d")

    # 确保字段存在
    show_cols = [
        "证券名称", "日期", "交易份额（万份）", "卖出主体", "买入主体",
        "交易净价", "交易全价", "交易金额（万元）"
    ]

    # 只显示存在的列
    exist_cols = [c for c in show_cols if c in trades_show.columns]

    # ---------------------
    # ✅ 汇总求和（交易份额 + 交易金额）
    # ---------------------
    total_share = trades_show["交易份额（万份）"].sum()
    total_amount = trades_show["交易金额（万元）"].sum()

    st.dataframe(
        trades_show[exist_cols].style.format(precision=2),
        width="stretch",
        height=450,
        hide_index=True
    )

    # ---------------------
    # ✅ 底部汇总行（卡片内）
    # ---------------------
    st.markdown(f"""
    <div style="margin-top:10px; font-size:15px; font-weight:bold; color:#1e40af;">
        合计： 交易份额 = {total_share:,.2f} 万份  |  交易金额 = {total_amount:,.2f} 万元
    </div>
    """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ----------------------
# 页面底部信息
# ----------------------
# 超紧凑分割线（间距极小）
st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)
st.markdown(f"""
<div style='text-align: center; color: #6b7280; font-size: 0.9rem;'>
    ABS投资数据统计分析平台 | 数据更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
</div>
""", unsafe_allow_html=True)
