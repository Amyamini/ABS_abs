"""
产品要素管理页面
记录每只产品的申购赎回到账时间等关键要素信息
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import os


def load_product_elements():
    """加载产品要素表"""
    file_path = "产品要素表.xlsx"
    
    if os.path.exists(file_path):
        try:
            df = pd.read_excel(file_path)
            # 确保所有列都是正确的数据类型
            if "备注" in df.columns:
                df["备注"] = df["备注"].fillna("").astype(str)
            if "申赎渠道" in df.columns:
                df["申赎渠道"] = df["申赎渠道"].fillna("").astype(str)
            if "产品名称" in df.columns:
                df["产品名称"] = df["产品名称"].fillna("").astype(str)
            # 确保数值列是数值类型
            for col in ["申购到账时间(T+N)", "赎回到账时间(T+N)"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            return df
        except Exception as e:
            st.error(f"加载产品要素表失败：{e}")
            return pd.DataFrame()
    else:
        # 如果文件不存在，创建空的DataFrame
        return pd.DataFrame(columns=[
            "产品名称", "申赎渠道", "申购到账时间(T+N)", 
            "赎回到账时间(T+N)", "备注"
        ])


def save_product_elements(df):
    """保存产品要素表"""
    try:
        df.to_excel("产品要素表.xlsx", index=False)
        return True
    except Exception as e:
        st.error(f"保存失败：{e}")
        return False


def render_product_elements():
    """渲染产品要素管理页面"""
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>
            📋 产品要素管理
        </div>
    """, unsafe_allow_html=True)
    
    st.caption("记录每只产品的申购赎回到账时间等关键要素信息")
    
    # 加载数据
    df_elements = load_product_elements()
    
    if df_elements.empty:
        st.info("暂无产品要素数据，请添加新产品")
        # 创建空的DataFrame用于编辑
        df_elements = pd.DataFrame(columns=[
            "产品名称", "申赎渠道", "申购到账时间(T+N)",
            "赎回到账时间(T+N)", "备注"
        ])
    else:
        # 确保数据类型正确
        df_elements = df_elements.copy()
        for col in ["备注", "申赎渠道", "产品名称"]:
            if col in df_elements.columns:
                df_elements[col] = df_elements[col].fillna("").astype(str)
        for col in ["申购到账时间(T+N)", "赎回到账时间(T+N)"]:
            if col in df_elements.columns:
                df_elements[col] = pd.to_numeric(df_elements[col], errors='coerce').fillna(0).astype(int)
    
    # ============================
    # 功能1：查看和编辑产品要素表
    # ============================
    st.subheader("📊 产品要素列表")
    
    # 搜索框
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        search_keyword = st.text_input("搜索产品名称", placeholder="输入关键词...")
    
    # 应用搜索
    df_display = df_elements.copy()
    if search_keyword:
        df_display = df_display[
            df_display["产品名称"].str.contains(search_keyword, na=False, case=False)
        ]
    
    # 确保显示的数据类型正确
    if not df_display.empty:
        for col in ["备注", "申赎渠道", "产品名称"]:
            if col in df_display.columns:
                df_display[col] = df_display[col].fillna("").astype(str)
    
    if df_display.empty and search_keyword:
        st.warning(f"未找到包含 '{search_keyword}' 的产品")
    elif df_display.empty and not search_keyword:
        # 数据为空且没有搜索，显示提示
        st.info("📝 暂无产品数据，请在下方表格中添加新产品")
        # 显示空表格供用户添加数据
        edited_df = st.data_editor(
            df_elements,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "产品名称": st.column_config.TextColumn("产品名称", width="medium"),
                "申赎渠道": st.column_config.TextColumn("申赎渠道", width="medium", help="如：直销、代销、银行等"),
                "申购到账时间(T+N)": st.column_config.NumberColumn(
                    "申购到账时间(T+N)",
                    min_value=0,
                    max_value=30,
                    step=1,
                    help="T+0表示当天到账，T+1表示下一个工作日到账",
                    width="small"
                ),
                "赎回到账时间(T+N)": st.column_config.NumberColumn(
                    "赎回到账时间(T+N)",
                    min_value=0,
                    max_value=30,
                    step=1,
                    help="T+0表示当天到账，T+7表示7个工作日后到账",
                    width="small"
                ),
                "备注": st.column_config.TextColumn("备注", width="large"),
            },
            hide_index=True
        )
        
        # 保存按钮
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 3])
        with col_btn1:
            if st.button("💾 保存修改", type="primary"):
                if save_product_elements(edited_df):
                    st.success("✅ 保存成功！")
                    st.rerun()
        
        with col_btn2:
            if st.button("🔄 刷新"):
                st.rerun()
    elif not df_display.empty:
        # 使用data_editor进行编辑
        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "产品名称": st.column_config.TextColumn("产品名称", width="medium"),
                "申赎渠道": st.column_config.TextColumn("申赎渠道", width="medium", help="如：直销、代销、银行等"),
                "申购到账时间(T+N)": st.column_config.NumberColumn(
                    "申购到账时间(T+N)",
                    min_value=0,
                    max_value=30,
                    step=1,
                    help="T+0表示当天到账，T+1表示下一个工作日到账",
                    width="small"
                ),
                "赎回到账时间(T+N)": st.column_config.NumberColumn(
                    "赎回到账时间(T+N)",
                    min_value=0,
                    max_value=30,
                    step=1,
                    help="T+0表示当天到账，T+7表示7个工作日后到账",
                    width="small"
                ),
                "备注": st.column_config.TextColumn("备注", width="large"),
            },
            hide_index=True
        )
        
        # 保存按钮
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 3])
        with col_btn1:
            if st.button("💾 保存修改", type="primary"):
                if save_product_elements(edited_df):
                    st.success("✅ 保存成功！")
                    st.rerun()
        
        with col_btn2:
            if st.button("🔄 刷新"):
                st.rerun()
    
    st.markdown("---")
    
    # ============================
    # 功能2：快速查询工具
    # ============================
    st.subheader("🔍 到账时间查询")
    
    if df_elements.empty:
        st.info("💡 请先在上方表格中添加产品数据，然后即可使用查询功能")
    else:
        col_query1, col_query2, col_query3 = st.columns(3)
        
        with col_query1:
            selected_product = st.selectbox(
                "选择产品",
                [""] + sorted(df_elements["产品名称"].dropna().unique().tolist()),
                help="选择要查询的产品"
            )
        
        if selected_product:
            product_info = df_elements[df_elements["产品名称"] == selected_product].iloc[0]
            
            with col_query2:
                st.metric(
                    "申购到账时间",
                    f"T+{int(product_info.get('申购到账时间(T+N)', 0))}",
                    help="申购资金到账所需工作日"
                )
            
            with col_query3:
                st.metric(
                    "赎回到账时间",
                    f"T+{int(product_info.get('赎回到账时间(T+N)', 0))}",
                    help="赎回资金到账所需工作日"
                )
            
            # 显示详细信息
            st.markdown(f"""
            <div style='background-color: #f0f9ff; padding: 15px; border-radius: 8px; margin-top: 10px;'>
                <h4 style='margin-top: 0; color: #1e40af;'>{selected_product} - 详细信息</h4>
                <p><strong>申赎渠道：</strong>{product_info.get('申赎渠道', '—')}</p>
                <p><strong>申购到账时间：</strong>T+{int(product_info.get('申购到账时间(T+N)', 0))} 工作日</p>
                <p><strong>赎回到账时间：</strong>T+{int(product_info.get('赎回到账时间(T+N)', 0))} 工作日</p>
                <p><strong>备注：</strong>{product_info.get('备注', '—')}</p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # ============================
    # 功能3：使用说明
    # ============================
    with st.expander("📖 使用说明"):
        st.markdown("""
        ### 字段说明
            
        - **产品名称**：产品的完整名称
        - **申赎渠道**：产品申购赎回的渠道（如：直销、代销、银行、券商等）
        - **申购到账时间(T+N)**：申购后资金确认到账所需的工作日数
          - T+0：当天到账
          - T+1：下一个工作日到账
          - T+2：两个工作日后到账
        - **赎回到账时间(T+N)**：赎回后资金到账所需的工作日数
          - 通常比申购时间长，因为需要清算时间
        - **备注**：其他需要说明的事项
            
        ### 操作提示
            
        1. **添加新产品**：在表格最后一行输入新产品信息
        2. **编辑现有产品**：直接点击单元格进行修改
        3. **删除产品**：选中行后按Delete键
        4. **保存修改**：点击“保存修改”按钮保存到Excel文件
        5. **查询功能**：使用搜索框或下拉框快速查找产品
            
        ### 应用场景
            
        - 流动性管理：根据到账时间安排资金调度
        - 交易决策：了解不同产品的流动性特征
        - 渠道管理：记录不同产品的申赎渠道信息
        """)


# ============================
# 独立运行入口
# ============================
if __name__ == "__main__":
    st.set_page_config(
        page_title="产品要素管理",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # 自定义CSS样式（与主应用保持一致）
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
    </style>
    """, unsafe_allow_html=True)
    
    # 渲染页面
    render_product_elements()
