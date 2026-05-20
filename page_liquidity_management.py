"""
产品流动性管理表
根据产品要素表的申购赎回到账时间，生成每只产品的现金流预测和管理表
"""
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os


def load_product_elements():
    """加载产品要素表（后台调用）"""
    file_path = "产品要素表.xlsx"
    
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path)
            # 确保数据类型正确
            if "备注" in df.columns:
                df["备注"] = df["备注"].fillna("").astype(str)
            if "申赎渠道" in df.columns:
                df["申赎渠道"] = df["申赎渠道"].fillna("").astype(str)
            if "产品名称" in df.columns:
                df["产品名称"] = df["产品名称"].fillna("").astype(str)
            for col in ["申购到账时间(T+N)", "赎回到账时间(T+N)"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            return df
        except Exception as e:
            st.error(f"加载产品要素表失败：{e}")
            return pd.DataFrame()
    else:
        return pd.DataFrame(columns=[
            "产品名称", "申赎渠道", "申购到账时间(T+N)", 
            "赎回到账时间(T+N)", "备注"
        ])


def generate_liquidity_table(product_name, start_date=None, end_date=None, initial_balance=0):
    """
    生成单只产品的流动性管理表
    
    Args:
        product_name: 产品名称
        start_date: 开始日期（默认今天）
        end_date: 结束日期（默认30天后）
        initial_balance: 初始余额
    
    Returns:
        DataFrame: 包含日期、现金流入、现金流出、期末余额
    """
    if start_date is None:
        start_date = datetime.now().date()
    if end_date is None:
        end_date = start_date + timedelta(days=30)
    
    # 加载产品要素
    df_elements = load_product_elements()
    
    # 查找产品信息
    if product_name not in df_elements["产品名称"].values:
        st.warning(f"未找到产品 '{product_name}' 的要素信息")
        return pd.DataFrame()
    
    product_info = df_elements[df_elements["产品名称"] == product_name].iloc[0]
    
    # 获取到账时间
    subscribe_days = int(product_info.get("申购到账时间(T+N)", 0))
    redeem_days = int(product_info.get("赎回到账时间(T+N)", 0))
    
    # 生成日期序列
    dates = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # 初始化数据
    data = []
    current_balance = initial_balance
    
    for date in dates:
        # 这里可以根据实际业务逻辑生成现金流
        # 目前生成示例数据：随机申购和赎回
        cash_inflow = 0
        cash_outflow = 0
        
        # 模拟：每隔几天有申购或赎回
        day_of_period = (date - dates[0]).days
        
        # 示例：每5天有一次申购，每7天有一次赎回
        if day_of_period % 5 == 0 and day_of_period > 0:
            cash_inflow = 1000000  # 100万申购
        
        if day_of_period % 7 == 0 and day_of_period > 0:
            cash_outflow = 500000  # 50万赎回
        
        # 计算期末余额
        current_balance = current_balance + cash_inflow - cash_outflow
        
        data.append({
            "日期": date.strftime("%Y-%m-%d"),
            "现金流入": cash_inflow,
            "现金流出": cash_outflow,
            "期末余额": current_balance
        })
    
    df_liquidity = pd.DataFrame(data)
    
    return df_liquidity


def save_liquidity_table(product_name, df_liquidity):
    """保存流动性管理表到Excel"""
    try:
        filename = f"流动性管理表_{product_name}.xlsx"
        df_liquidity.to_excel(filename, index=False)
        return True, filename
    except Exception as e:
        return False, str(e)


def render_liquidity_management():
    """渲染流动性管理页面"""
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>
            💰 产品流动性管理表
        </div>
    """, unsafe_allow_html=True)
    
    st.caption("根据产品要素表的申购赎回到账时间，生成每只产品的现金流预测和管理表")
    
    # 加载产品要素（后台）
    df_elements = load_product_elements()
    
    if df_elements.empty:
        st.warning("⚠️ 未找到产品要素数据，请先在产品要素表中添加产品信息")
        st.info("💡 提示：产品要素表用于后台配置，不在前台展示")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    
    # ============================
    # 产品选择
    # ============================
    st.subheader("📋 选择产品")
    
    product_list = sorted(df_elements["产品名称"].dropna().unique().tolist())
    
    if not product_list:
        st.warning("暂无产品数据")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    
    selected_product = st.selectbox(
        "选择要查看的产品",
        product_list,
        help="选择要生成流动性管理表的产品"
    )
    
    # ============================
    # 参数设置
    # ============================
    st.subheader("⚙️ 参数设置")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_date = st.date_input(
            "开始日期",
            value=datetime.now().date(),
            help="流动性管理表的起始日期"
        )
    
    with col2:
        days_range = st.number_input(
            "天数范围",
            min_value=7,
            max_value=365,
            value=30,
            step=7,
            help="生成多少天的流动性预测"
        )
        end_date = start_date + timedelta(days=days_range)
    
    with col3:
        initial_balance = st.number_input(
            "初始余额（元）",
            min_value=0,
            value=10000000,
            step=1000000,
            help="期初的现金余额"
        )
    
    # 显示产品要素信息（简要）
    if selected_product:
        product_info = df_elements[df_elements["产品名称"] == selected_product].iloc[0]
        subscribe_days = int(product_info.get("申购到账时间(T+N)", 0))
        redeem_days = int(product_info.get("赎回到账时间(T+N)", 0))
        
        st.info(f"""
        📊 **{selected_product}** 的产品要素：
        - 申购到账时间：T+{subscribe_days}
        - 赎回到账时间：T+{redeem_days}
        - 申赎渠道：{product_info.get('申赎渠道', '—')}
        """)
    
    # ============================
    # 生成流动性管理表
    # ============================
    if st.button("🔄 生成流动性管理表", type="primary"):
        with st.spinner("正在生成流动性管理表..."):
            df_liquidity = generate_liquidity_table(
                selected_product,
                start_date,
                end_date,
                initial_balance
            )
            
            if not df_liquidity.empty:
                # 保存到session_state
                st.session_state[f"liquidity_{selected_product}"] = df_liquidity
                
                st.success(f"✅ 已生成 {selected_product} 的流动性管理表（{len(df_liquidity)}天）")
    
    # ============================
    # 显示流动性管理表
    # ============================
    liquidity_key = f"liquidity_{selected_product}"
    
    if liquidity_key in st.session_state:
        df_liquidity = st.session_state[liquidity_key]
        
        st.subheader(f"📊 {selected_product} - 流动性管理表")
        
        # 格式化显示
        df_display = df_liquidity.copy()
        
        # 使用data_editor显示（可编辑）
        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "日期": st.column_config.TextColumn("日期", width="small"),
                "现金流入": st.column_config.NumberColumn(
                    "现金流入（元）",
                    format="%,.2f",
                    width="medium"
                ),
                "现金流出": st.column_config.NumberColumn(
                    "现金流出（元）",
                    format="%,.2f",
                    width="medium"
                ),
                "期末余额": st.column_config.NumberColumn(
                    "期末余额（元）",
                    format="%,.2f",
                    width="medium"
                ),
            },
            hide_index=True
        )
        
        # 统计信息
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        
        with col_stat1:
            total_inflow = edited_df["现金流入"].sum()
            st.metric("总现金流入", f"¥{total_inflow:,.2f}")
        
        with col_stat2:
            total_outflow = edited_df["现金流出"].sum()
            st.metric("总现金流出", f"¥{total_outflow:,.2f}")
        
        with col_stat3:
            net_flow = total_inflow - total_outflow
            st.metric("净现金流", f"¥{net_flow:,.2f}")
        
        with col_stat4:
            final_balance = edited_df["期末余额"].iloc[-1] if len(edited_df) > 0 else 0
            st.metric("期末余额", f"¥{final_balance:,.2f}")
        
        # 保存按钮
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            if st.button("💾 保存到Excel"):
                success, result = save_liquidity_table(selected_product, edited_df)
                if success:
                    st.success(f"✅ 已保存到文件：{result}")
                else:
                    st.error(f"❌ 保存失败：{result}")
        
        with col_btn2:
            if st.button("🗑️ 清除数据"):
                if liquidity_key in st.session_state:
                    del st.session_state[liquidity_key]
                    st.rerun()
    
    else:
        st.info("👆 请点击'生成流动性管理表'按钮生成数据")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # ============================
    # 使用说明
    # ============================
    with st.expander("📖 使用说明"):
        st.markdown("""
        ### 功能说明
        
        本页面根据**产品要素表**中配置的申购赎回到账时间，为每只产品生成流动性管理表。
        
        ### 操作流程
        
        1. **选择产品**：从下拉列表中选择要管理的产品
        2. **设置参数**：
           - 开始日期：流动性管理的起始日期
           - 天数范围：生成多少天的预测（7-365天）
           - 初始余额：期初的现金余额
        3. **生成表格**：点击"生成流动性管理表"按钮
        4. **编辑数据**：可以直接在表格中修改现金流数据
        5. **保存导出**：点击"保存到Excel"导出为Excel文件
        
        ### 字段说明
        
        - **日期**：每一天的日期
        - **现金流入**：当天的现金流入金额（如申购款到账）
        - **现金流出**：当天的现金流出金额（如赎回款支付）
        - **期末余额**：当天结束时的现金余额
        
        ### 注意事项
        
        - 产品要素表作为后台配置，不在本页面展示
        - 当前生成的是示例数据，实际使用时需要根据真实交易记录生成
        - 可以手动编辑表格中的现金流数据
        - 保存的Excel文件命名格式：`流动性管理表_产品名称.xlsx`
        
        ### 后续优化
        
        - 接入真实交易数据，自动生成现金流
        - 根据申购赎回到账时间自动计算资金到账日期
        - 支持批量生成多个产品的流动性管理表
        - 添加现金流图表可视化
        """)


# ============================
# 独立运行入口
# ============================
if __name__ == "__main__":
    st.set_page_config(
        page_title="产品流动性管理",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # 自定义CSS样式
    st.markdown("""
    <style>
        .block-container {
            padding-top: 1rem !important;
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
        .metric-card {
            background-color: #f0f9ff;
            border-radius: 8px;
            padding: 15px;
            margin: 5px 0;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # 渲染页面
    render_liquidity_management()
