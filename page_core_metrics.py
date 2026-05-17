"""
核心投资指标页面模块
展示核心投资数据和图表分析
"""
import streamlit as st
import pandas as pd
import plotly.express as px


def render_core_metrics(stats, df_filtered, trades_filtered):
    """
    渲染核心投资指标页面
    
    Args:
        stats: 统计指标字典
        df_filtered: 过滤后的主数据
        trades_filtered: 过滤后的交易记录
    """
    # st.subheader("核心投资指标")
    col1, col2, col3, col4 = st.columns(4, gap="small")

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
                yaxis=dict(
                    automargin=True,
                    fixedrange=False
                )
            )

            # 核心：默认只显示【近12个月】
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
                x="年份",
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
            st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.warning("缺少 日期/产品名称/本金-现金流 字段，无法生成图表")

    # 按产品分组，计算 本金-现金流 总和（投资本金余额）
    if "产品名称" in df_filtered.columns and "本金-现金流" in df_filtered.columns and "日期" in df_filtered.columns:
        product_balance = df_filtered.groupby("产品名称").agg(
            投资本金余额=("本金-现金流", "sum"),
            最早投资日期=("日期", "min")
        ).reset_index()
        product_balance.columns = ["产品名称", "投资本金余额", "最早投资日期"]

        # 只保留 投资本金余额 < 0 的产品
        product_balance = product_balance[product_balance["投资本金余额"] < 0]

        # 转为正数（更符合规模展示）
        product_balance["投资本金余额（万元）"] = product_balance["投资本金余额"].round(2).abs() / 10000

        # 计算占比
        total_balance = product_balance["投资本金余额（万元）"].sum()
        product_balance["占比(%)"] = (product_balance["投资本金余额（万元）"] / total_balance * 100).round(2)

        # 按金额从大到小排序
        product_balance = product_balance.sort_values("投资本金余额（万元）", ascending=False)

        # 超紧凑分割线（间距极小）
        st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)
        st.markdown("""
        <div class='card'>
            <div class='card-title' style='color: black;'>
                产品存续规模Top20
            </div>
        """, unsafe_allow_html=True)
        
        product_balance_top20 = product_balance.head(20)
        fig = px.bar(
            product_balance_top20,
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
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.warning("暂无产品规模数据")

    st.markdown("---")
