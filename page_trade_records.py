"""
交易记录页面模块
展示交易记录明细和筛选功能
"""
import streamlit as st
import pandas as pd


def render_trade_records(trades_filtered):
    """
    渲染交易记录页面
    
    Args:
        trades_filtered: 过滤后的交易记录数据
    """
    # 调试信息：显示可用的列名
    # st.info(f"交易记录可用列：{', '.join(trades_filtered.columns.tolist())}")
    
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
        securities_local = st.text_input("证券名称", placeholder="模糊搜索...")

    with filter_col3:
        sellers = st.text_input("卖出方", placeholder="模糊搜索...")

    with filter_col4:
        buyers = st.text_input("买入方", placeholder="模糊搜索...")

    st.markdown("</div>", unsafe_allow_html=True)

    # 应用筛选
    trades_show = trades_filtered.copy()
    
    # 显示当前数据量（用于调试）
    st.caption(f"📊 当前显示 {len(trades_show)} 条交易记录（已应用全局筛选）")

    # 时间筛选
    trades_show = trades_show[
        (trades_show["日期"] >= pd.to_datetime(date_range[0])) &
        (trades_show["日期"] <= pd.to_datetime(date_range[1]))
    ]

    # 证券筛选（局部筛选，与全局筛选叠加）
    if securities_local:
        trades_show = trades_show[trades_show["证券名称"].str.contains(securities_local, na=False, case=False)]

    # 买入方筛选
    if buyers:
        trades_show = trades_show[trades_show["买入主体"].str.contains(buyers, na=False, case=False)]

    # 卖出方筛选
    if sellers:
        trades_show = trades_show[trades_show["卖出主体"].str.contains(sellers, na=False, case=False)]

    # 格式化日期为yyyy-mm-dd
    trades_show["日期"] = trades_show["日期"].dt.strftime("%Y-%m-%d")

    # 确保字段存在
    show_cols = [
        "证券名称", "日期", "交易份额（万份）", "卖出主体", "买入主体",
        "交易净价", "交易全价", "交易金额（元）"
    ]

    # 只显示存在的列
    exist_cols = [c for c in show_cols if c in trades_show.columns]

    # ---------------------
    # ✅ 汇总求和（交易份额 + 交易金额）
    # ---------------------
    total_share = trades_show["交易份额（万份）"].sum()
    total_amount = trades_show["交易金额（元）"].sum()

    # 格式化表格数据
    formatter = {
        "交易份额（万份）": "{:.0f}",
        "交易净价": "{:.4f}",
        "交易全价": "{:.4f}",
        "交易金额（元）": "{:,.2f}"
    }
    # 只格式化存在的列
    exist_formatter = {k: v for k, v in formatter.items() if k in exist_cols}

    st.dataframe(
        trades_show[exist_cols].style.format(exist_formatter),
        width="stretch",
        height=450,
        hide_index=True
    )

    # ---------------------
    # ✅ 底部汇总行（卡片内）
    # ---------------------
    st.markdown(f"""
    <div style="margin-top:10px; font-size:15px; font-weight:bold; color:#1e40af;">
        合计： 交易份额 = {total_share:,.2f} 万份   |   交易金额 = {total_amount:,.2f} 元
    </div>
    """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
