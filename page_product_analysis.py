"""
产品分析页面模块
展示产品持仓及收益明细
"""
import streamlit as st
import pandas as pd


def render_product_analysis(df_filtered):
    """
    渲染产品分析页面
    
    Args:
        df_filtered: 过滤后的主数据
    """
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
        "利息-现金流", "本金-现金流", "投资现金流"
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
        with col2:
            security_list = st.text_input("证券名称", placeholder="输入证券关键词...")

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
                import scipy.optimize as sco

                def xirr(cashflows, dates, guess=0.1):
                    try:
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
                合计： 持仓份额 = {total_share_final_table:,.2f} 万份   |   投资本金余额 = {total_amount_final_table:,.2f} 万元   |   收益金额 = {total_profit_final_table:,.2f} 万元
            </div>
            """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)
