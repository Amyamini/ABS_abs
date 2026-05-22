import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, date
import io
import os

st.set_page_config(
    layout="wide", 
    page_title="产品流动性管理系统",
    initial_sidebar_state="collapsed"
)

# ----------------------
# 产品参数加载函数
# ----------------------
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
            
            # 处理简称列（可能有多种命名方式）
            alias_col = None
            for col_name in ["简称", "产品简称", "缩写", "产品缩写"]:
                if col_name in df.columns:
                    alias_col = col_name
                    break
            
            if alias_col:
                df[alias_col] = df[alias_col].fillna("").astype(str)
            
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


def create_product_name_mapping(df_elements):
    """创建产品名称映射：简称 -> 全称"""
    if df_elements.empty or "产品名称" not in df_elements.columns:
        return {}
    
    mapping = {}
    
    # 查找简称列
    alias_col = None
    for col_name in ["简称", "产品简称", "缩写", "产品缩写"]:
        if col_name in df_elements.columns:
            alias_col = col_name
            break
    
    if alias_col:
        # 创建简称到全称的映射
        for _, row in df_elements.iterrows():
            full_name = row["产品名称"]
            alias = row[alias_col]
            if alias and str(alias).strip():
                mapping[str(alias).strip()] = full_name
            # 也添加全称到自身的映射
            mapping[full_name] = full_name
    else:
        # 如果没有简称列，只使用产品名称
        for _, row in df_elements.iterrows():
            full_name = row["产品名称"]
            mapping[full_name] = full_name
    
    return mapping


def save_product_elements(df):
    """保存产品要素表"""
    try:
        df.to_excel("产品要素表.xlsx", index=False)
        return True
    except Exception as e:
        st.error(f"保存失败：{e}")
        return False


# ----------------------
# 主界面：流动性管理表
# ----------------------

# 自定义CSS优化页面布局
st.markdown("""
<style>
    /* 减小标题和文本的间距 */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0.5rem;
    }
    
    /* 减小表格行高 */
    .stDataFrame [data-testid="stVerticalBlock"] {
        gap: 0.2rem;
    }
    
    /* 优化按钮样式 */
    .stButton > button {
        padding: 0.25rem 0.5rem;
        font-size: 0.85rem;
    }
    
    /* 减小选择框高度 */
    .stSelectbox > div > div {
        min-height: 2.5rem;
    }
    
    /* 优化表格字体大小 */
    div[data-testid="stDataframe"] {
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏦 产品流动性管理系统")

# ---------- 1. 加载数据 ----------
DEFAULT_FILE = "基金流动性管理总表.xlsx"
file_path = DEFAULT_FILE

# 检查文件是否存在
if not os.path.exists(file_path):
    st.error(f"❌ 找不到文件：{file_path}")
    st.stop()


@st.cache_data(ttl=300)  # 缓存5分钟，减少重新加载
def load_all_sheets(file_path):
    """优化版：使用 xls 对象直接读取，避免重复打开文件"""
    # 一次性创建 ExcelFile 对象
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheets_data = {}
    
    # 只读取从第3张sheet开始（索引2）的所有sheet
    target_sheets = xls.sheet_names[2:]  # 从索引2开始，即第3张sheet
    
    for sheet_name in target_sheets:
        # 使用 xls.parse() 而不是 pd.read_excel()，避免重复读取文件
        df = xls.parse(sheet_name)
        
        # 取前6列，标准化列名（新增“关联产品”列）
        if df.shape[1] >= 6:
            flow_df = df.iloc[:, :6].copy()
            flow_df.columns = ["日期", "现金流类型", "关联产品", "现金流入", "现金流出", "期末余额"]
        elif df.shape[1] >= 5:
            flow_df = df.iloc[:, :5].copy()
            flow_df.columns = ["日期", "现金流类型", "现金流入", "现金流出", "期末余额"]
            # 插入空的“关联产品”列
            flow_df.insert(2, "关联产品", "")
        else:
            flow_df = df.copy()
            # 补充缺失的列
            for i in range(6 - df.shape[1]):
                flow_df[f"col_{i + df.shape[1] + 1}"] = ""
            flow_df.columns = ["日期", "现金流类型", "关联产品", "现金流入", "现金流出", "期末余额"] + list(flow_df.columns[6:])
            flow_df = flow_df.iloc[:, :6]
        
        # 数值转换 - 使用向量化操作提高性能
        for col in ["现金流入", "现金流出", "期末余额"]:
            flow_df[col] = pd.to_numeric(flow_df[col], errors="coerce")
        # 日期转换 - 指定常见日期格式以避免警告
        flow_df["日期"] = pd.to_datetime(flow_df["日期"], errors="coerce", format="mixed", dayfirst=False)
        
        # 确保“关联产品”列是字符串类型
        flow_df["关联产品"] = flow_df["关联产品"].fillna("").astype(str)
        
        sheets_data[sheet_name] = flow_df
    
    return sheets_data


try:
    sheets_data = load_all_sheets(file_path)
    all_products = list(sheets_data.keys())
except Exception as e:
    st.error(f"读取文件失败：{e}")
    st.stop()

# 加载产品要素表并创建名称映射
df_elements = load_product_elements()
name_mapping = create_product_name_mapping(df_elements)

# 创建反向映射：全称 -> 简称（用于显示）
reverse_mapping = {}
if df_elements is not None and not df_elements.empty:
    alias_col = None
    for col_name in ["简称", "产品简称", "缩写", "产品缩写"]:
        if col_name in df_elements.columns:
            alias_col = col_name
            break
    
    if alias_col:
        for _, row in df_elements.iterrows():
            full_name = row["产品名称"]
            alias = row[alias_col]
            if alias and str(alias).strip():
                reverse_mapping[full_name] = str(alias).strip()

# ---------- 2. 产品类型配置（存储在 session_state） ----------
if "product_types" not in st.session_state:
    # 初始默认识别：优先从产品要素表读取，其次根据工作表名称关键词
    st.session_state.product_types = {}
    
    # 查找类型列（可能有多种命名方式）
    type_col = None
    if not df_elements.empty:
        for col_name in ["产品类型", "类型", "基金类型"]:
            if col_name in df_elements.columns:
                type_col = col_name
                break
    
    for p in all_products:
        # 1. 优先从产品要素表读取类型
        if type_col and not df_elements.empty:
            # 尝试通过名称匹配
            if p in df_elements["产品名称"].values:
                product_type = df_elements[df_elements["产品名称"] == p].iloc[0][type_col]
                st.session_state.product_types[p] = str(product_type) if pd.notna(product_type) else "子基金"
            # 尝试通过简称匹配
            else:
                alias_col = None
                for col_name in ["简称", "产品简称", "缩写", "产品缩写"]:
                    if col_name in df_elements.columns:
                        alias_col = col_name
                        break
                
                if alias_col and p in df_elements[alias_col].values:
                    product_type = df_elements[df_elements[alias_col] == p].iloc[0][type_col]
                    st.session_state.product_types[p] = str(product_type) if pd.notna(product_type) else "子基金"
                else:
                    # 2. 如果产品要素表没有，根据关键词判断
                    if "母基金" in p or "FOF" in p or "FOF" in p.upper():
                        st.session_state.product_types[p] = "母基金"
                    else:
                        st.session_state.product_types[p] = "子基金"
        else:
            # 3. 如果没有产品要素表，根据关键词判断
            if "母基金" in p or "FOF" in p or "FOF" in p.upper():
                st.session_state.product_types[p] = "母基金"
            else:
                st.session_state.product_types[p] = "子基金"

# ---------- 3. 辅助函数：保存单个 sheet 到 Excel ----------
def save_sheet_to_excel(sheet_name, df_updated, wb_path):
    wb = load_workbook(wb_path)
    ws = wb[sheet_name]
    # 清空前6列所有行（从第1行开始）
    for row in range(1, ws.max_row + 1):
        for col in range(1, 7):  # 改为7列
            ws.cell(row=row, column=col, value=None)
    # 写入新数据（包含表头）
    for r_idx, row in enumerate(dataframe_to_rows(df_updated, index=False, header=True), 1):
        for c_idx, value in enumerate(row[:6], 1):  # 改为6列
            ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(wb_path)


def add_cashflow_record(product_name, date_val, flow_type, related_product, inflow, outflow, balance_after):
    """向指定产品的现金流表追加一行，并返回新的 DataFrame"""
    df = sheets_data[product_name].copy()
    new_row = pd.DataFrame({
        "日期": [date_val],
        "现金流类型": [flow_type],
        "关联产品": [related_product],
        "现金流入": [inflow],
        "现金流出": [outflow],
        "期末余额": [balance_after]
    })
    df = pd.concat([df, new_row], ignore_index=True)
    # 确保数值类型
    for col in ["现金流入", "现金流出", "期末余额"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    sheets_data[product_name] = df
    return df


# ---------- 4. 主界面：多产品选择与展示 ----------
st.header("📋 产品现金流明细（可编辑）")

# 多产品选择器 - 更紧凑的布局
st.markdown("**选择要查看的产品（最多3个）**")
col_select1, col_select2, col_select3 = st.columns(3)

with col_select1:
    selected_product1 = st.selectbox("产品 1", [""] + all_products, key="product_select_1", label_visibility="collapsed")
with col_select2:
    selected_product2 = st.selectbox("产品 2", [""] + all_products, key="product_select_2", label_visibility="collapsed")
with col_select3:
    selected_product3 = st.selectbox("产品 3", [""] + all_products, key="product_select_3", label_visibility="collapsed")

# 收集已选择的产品
selected_products = [p for p in [selected_product1, selected_product2, selected_product3] if p]

if not selected_products:
    st.info("👈 请从上方下拉框中选择要查看的产品")
else:
    # 检查是否点击了某个产品进入参数页面
    if st.session_state.get('viewing_product_params', None):
        viewing_product = st.session_state['viewing_product_params']
        
        # 显示产品参数页面
        st.markdown(f"# 📋 {viewing_product} - 产品参数")
        
        # 返回按钮
        if st.button("⬅️ 返回现金流管理", type="primary"):
            st.session_state['viewing_product_params'] = None
            st.rerun()
        
        st.markdown("---")
        
        # 加载产品要素表
        df_elements = load_product_elements()
        
        # 使用简称映射查找产品
        product_row = None
        
        if not df_elements.empty:
            # 查找简称列
            alias_col = None
            for col_name in ["简称", "产品简称", "缩写", "产品缩写"]:
                if col_name in df_elements.columns:
                    alias_col = col_name
                    break
            
            # 1. 首先尝试精确匹配产品名称
            if viewing_product in df_elements["产品名称"].values:
                product_row = df_elements[df_elements["产品名称"] == viewing_product].index[0]
            # 2. 如果有简称列，尝试通过简称匹配
            elif alias_col and viewing_product in df_elements[alias_col].values:
                product_row = df_elements[df_elements[alias_col] == viewing_product].index[0]
            # 3. 尝试通过映射查找全称
            elif name_mapping and viewing_product in name_mapping:
                full_name = name_mapping[viewing_product]
                if full_name in df_elements["产品名称"].values:
                    product_row = df_elements[df_elements["产品名称"] == full_name].index[0]
        
        if product_row is not None:
            product_info = df_elements.loc[product_row]
            
            # 显示并编辑产品参数
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                new_channel = st.text_input("申赎渠道", value=str(product_info.get("申赎渠道", "")), key=f"edit_channel_{viewing_product}")
                new_purchase_days = st.number_input("申购到账时间(T+N)", min_value=0, max_value=30, 
                                                    value=int(product_info.get("申购到账时间(T+N)", 0)), key=f"edit_purchase_{viewing_product}")
            with col_p2:
                new_redeem_days = st.number_input("赎回到账时间(T+N)", min_value=0, max_value=30,
                                                  value=int(product_info.get("赎回到账时间(T+N)", 0)), key=f"edit_redeem_{viewing_product}")
                new_notes = st.text_input("备注", value=str(product_info.get("备注", "")), key=f"edit_notes_{viewing_product}")
            
            col_save1, col_save2 = st.columns([1, 3])
            with col_save1:
                if st.button("💾 保存参数", type="primary", key=f"save_product_params_{viewing_product}"):
                    # 更新数据
                    df_elements.loc[product_row, "申赎渠道"] = new_channel
                    df_elements.loc[product_row, "申购到账时间(T+N)"] = new_purchase_days
                    df_elements.loc[product_row, "赎回到账时间(T+N)"] = new_redeem_days
                    df_elements.loc[product_row, "备注"] = new_notes
                    
                    if save_product_elements(df_elements):
                        st.success("✅ 产品参数保存成功！")
                        st.rerun()
            with col_save2:
                if st.button("🔄 取消编辑", key=f"cancel_edit_{viewing_product}"):
                    st.session_state['viewing_product_params'] = None
                    st.rerun()
        else:
            st.info(f"⚠️ 未找到「{viewing_product}」的参数信息，请在产品要素表中添加")
            
            # 提供快速添加功能
            col_add1, col_add2 = st.columns(2)
            with col_add1:
                add_channel = st.text_input("申赎渠道", placeholder="如：直销、代销", key=f"add_channel_{viewing_product}")
                add_purchase_days = st.number_input("申购到账时间(T+N)", min_value=0, max_value=30, value=0, key=f"add_purchase_{viewing_product}")
            with col_add2:
                add_redeem_days = st.number_input("赎回到账时间(T+N)", min_value=0, max_value=30, value=0, key=f"add_redeem_{viewing_product}")
                add_notes = st.text_input("备注", key=f"add_notes_{viewing_product}")
            
            if st.button("➕ 添加产品参数", type="primary", key=f"add_product_params_{viewing_product}"):
                new_row = pd.DataFrame({
                    "产品名称": [viewing_product],
                    "申赎渠道": [add_channel],
                    "申购到账时间(T+N)": [add_purchase_days],
                    "赎回到账时间(T+N)": [add_redeem_days],
                    "备注": [add_notes]
                })
                df_elements = pd.concat([df_elements, new_row], ignore_index=True)
                if save_product_elements(df_elements):
                    st.success("✅ 产品参数添加成功！")
                    st.rerun()
    
    else:
        # 创建三列布局，并排显示选中的产品
        product_columns = st.columns(len(selected_products))
        
        # 为每个选中的产品显示数据
        for idx, selected_product in enumerate(selected_products):
            with product_columns[idx]:
                # 显示产品类型和可点击的产品名称
                product_type = st.session_state.product_types[selected_product]
                
                # 使用链接样式的按钮作为产品名称
                if st.button(f"{idx + 1}. {selected_product}", key=f"product_link_{selected_product}", use_container_width=True):
                    st.session_state['viewing_product_params'] = selected_product
                    st.rerun()
                
                st.caption(f"类型：**{product_type}** | 💡 点击名称查看参数")
                
                # 显示现金流数据
                df = sheets_data[selected_product].copy()
                
                # 过滤：只显示今天往前推5天开始的数据
                from datetime import timedelta
                cutoff_date = pd.Timestamp.today() - timedelta(days=5)
                df = df[df["日期"] >= cutoff_date].copy()
                
                # 重置索引，避免显示原始行号
                df = df.reset_index(drop=True)
                
                # 填充NaN
                df["现金流入"] = df["现金流入"].fillna(0).astype(float)
                df["现金流出"] = df["现金流出"].fillna(0).astype(float)
                df["期末余额"] = df["期末余额"].fillna(0).astype(float)
                
                # 将关联产品列的全称转换为简称（用于显示）
                if reverse_mapping:
                    df["关联产品"] = df["关联产品"].apply(
                        lambda x: reverse_mapping.get(str(x).strip(), x) if str(x).strip() else ""
                    )
                
                # 创建简称列表用于下拉选择
                display_options = [""]  # 空选项
                for product in all_products:
                    # 如果有简称显示简称，否则显示全称
                    display_name = reverse_mapping.get(product, product)
                    display_options.append(display_name)
                
                edited_df = st.data_editor(
                    df,
                    use_container_width=True,
                    num_rows="dynamic",
                    hide_index=True,  # 隐藏索引列（序号列）
                    column_config={
                        "日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"),
                        "现金流类型": st.column_config.TextColumn("类型"),
                        "关联产品": st.column_config.SelectboxColumn(
                            "关联产品",
                            options=display_options,  # 使用简称列表
                            required=False
                        ),
                        "现金流入": st.column_config.NumberColumn("流入", step=1000),
                        "现金流出": st.column_config.NumberColumn("流出", step=1000),
                        "期末余额": st.column_config.NumberColumn("余额", step=1000),
                    },
                    key=f"editor_{selected_product}"
                )
                
                # 保存当前产品修改
                if st.button(f"💾 保存", type="primary", key=f"save_{selected_product}", use_container_width=True):
                    try:
                        # 将简称转换回全称（用于保存到Excel）
                        if name_mapping:
                            edited_df["关联产品"] = edited_df["关联产品"].apply(
                                lambda x: name_mapping.get(str(x).strip(), x) if str(x).strip() else ""
                            )
                        
                        wb_path = DEFAULT_FILE
                        save_sheet_to_excel(selected_product, edited_df, wb_path)
                        sheets_data[selected_product] = edited_df
                        st.success("✅ 保存成功！")
                        st.cache_data.clear()
                    except Exception as e:
                        st.error(f"❌ 保存失败：{e}")

st.markdown("---")
st.caption("💡 提示：可以同时选择最多3个产品进行对比查看，点击「查看/修改产品参数」按钮可以查看和编辑产品参数")