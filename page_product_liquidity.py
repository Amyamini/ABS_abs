import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, date
import io

st.set_page_config(layout="wide", page_title="产品流动性管理系统（母子基金联动）")
st.title("🏦 产品流动性管理系统（母子基金联动）")

# ---------- 1. 加载数据 ----------
uploaded_file = st.file_uploader("上传 Excel 文件（基金流动性管理总表.xlsx）", type=["xlsx"])
DEFAULT_FILE = "基金流动性管理总表.xlsx"

if uploaded_file is not None:
    file_path = uploaded_file
    file_name = uploaded_file.name
else:
    file_path = DEFAULT_FILE
    file_name = DEFAULT_FILE
    st.info(f"使用默认文件：{file_name}")


@st.cache_data(ttl=60)
def load_all_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheets_data = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet, engine="openpyxl")
        # 取前5列，标准化列名
        if df.shape[1] >= 5:
            flow_df = df.iloc[:, :5].copy()
            flow_df.columns = ["日期", "现金流类型", "现金流入", "现金流出", "期末余额"]
        else:
            flow_df = df.copy()
            for i in range(5 - df.shape[1]):
                flow_df[f"col_{i + df.shape[1] + 1}"] = ""
            flow_df.columns = ["日期", "现金流类型", "现金流入", "现金流出", "期末余额"] + list(flow_df.columns[5:])
            flow_df = flow_df.iloc[:, :5]
        # 数值转换
        for col in ["现金流入", "现金流出", "期末余额"]:
            flow_df[col] = pd.to_numeric(flow_df[col], errors="coerce")
        flow_df["日期"] = pd.to_datetime(flow_df["日期"], errors="coerce")
        sheets_data[sheet] = flow_df
    return sheets_data


try:
    sheets_data = load_all_sheets(file_path)
    all_products = list(sheets_data.keys())
except Exception as e:
    st.error(f"读取文件失败：{e}")
    st.stop()

# ---------- 2. 产品类型配置（存储在 session_state） ----------
if "product_types" not in st.session_state:
    # 初始默认识别：根据工作表名称包含“母基金”等关键词，没有则默认子基金
    st.session_state.product_types = {}
    for p in all_products:
        if "母基金" in p or "FOF" in p:
            st.session_state.product_types[p] = "母基金"
        else:
            st.session_state.product_types[p] = "子基金"

if "holdings" not in st.session_state:
    # 持仓结构：母基金 -> {子基金名称: 持有份额（金额，净值暂定为1）}
    st.session_state.holdings = {}

# 侧边栏配置产品类型
st.sidebar.header("产品类型配置")
for prod in all_products:
    current_type = st.session_state.product_types.get(prod, "子基金")
    new_type = st.sidebar.selectbox(
        f"{prod} 类型",
        ["母基金", "子基金"],
        index=0 if current_type == "母基金" else 1,
        key=f"type_{prod}"
    )
    st.session_state.product_types[prod] = new_type

# 更新持仓结构：确保每个母基金有一个持仓字典
for prod, ptype in st.session_state.product_types.items():
    if ptype == "母基金" and prod not in st.session_state.holdings:
        st.session_state.holdings[prod] = {}
    elif ptype == "子基金" and prod in st.session_state.holdings:
        # 如果某产品从母基金改为子基金，清除其持仓记录（可选）
        st.session_state.holdings.pop(prod, None)


# ---------- 3. 辅助函数：保存单个 sheet 的前5列到 Excel ----------
def save_sheet_to_excel(sheet_name, df_updated, wb_path):
    wb = load_workbook(wb_path)
    ws = wb[sheet_name]
    # 清空前5列所有行（从第1行开始）
    for row in range(1, ws.max_row + 1):
        for col in range(1, 6):
            ws.cell(row=row, column=col, value=None)
    # 写入新数据（包含表头）
    for r_idx, row in enumerate(dataframe_to_rows(df_updated, index=False, header=True), 1):
        for c_idx, value in enumerate(row[:5], 1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(wb_path)


def add_cashflow_record(product_name, date_val, flow_type, inflow, outflow, balance_after):
    """向指定产品的现金流表追加一行，并返回新的 DataFrame"""
    df = sheets_data[product_name].copy()
    new_row = pd.DataFrame({
        "日期": [date_val],
        "现金流类型": [flow_type],
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


# ---------- 4. 主界面：产品选择与展示 ----------
st.header("📋 产品现金流明细（可编辑）")
selected_product = st.selectbox("选择产品", all_products, key="product_select")
if selected_product:
    product_type = st.session_state.product_types[selected_product]
    st.caption(f"当前类型：**{product_type}**")

    df = sheets_data[selected_product].copy()
    # 填充NaN
    df["现金流入"] = df["现金流入"].fillna(0).astype(float)
    df["现金流出"] = df["现金流出"].fillna(0).astype(float)
    df["期末余额"] = df["期末余额"].fillna(0).astype(float)

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"),
            "现金流类型": st.column_config.TextColumn("现金流类型"),
            "现金流入": st.column_config.NumberColumn("现金流入 (元)", step=1000),
            "现金流出": st.column_config.NumberColumn("现金流出 (元)", step=1000),
            "期末余额": st.column_config.NumberColumn("期末余额 (元)", step=1000),
        }
    )

    # 保存当前产品修改
    if st.button(f"保存「{selected_product}」的修改"):
        try:
            if uploaded_file is not None:
                with open("temp_workbook.xlsx", "wb") as f:
                    f.write(uploaded_file.getbuffer())
                wb_path = "temp_workbook.xlsx"
            else:
                wb_path = DEFAULT_FILE
            save_sheet_to_excel(selected_product, edited_df, wb_path)
            sheets_data[selected_product] = edited_df
            st.success("保存成功！")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"保存失败：{e}")

# ---------- 5. 母子基金联动操作（母基金持仓管理） ----------
st.header("🔄 母子基金申购/赎回联动")
# 筛选出母基金列表
mother_funds = [p for p, t in st.session_state.product_types.items() if t == "母基金"]
if not mother_funds:
    st.info("尚未配置母基金产品，请在侧边栏将至少一个产品设为「母基金」")
else:
    selected_mother = st.selectbox("选择母基金", mother_funds, key="mother_select")
    if selected_mother:
        # 可投资的子基金列表（所有子基金产品）
        children = [p for p, t in st.session_state.product_types.items() if t == "子基金"]
        if not children:
            st.warning("没有子基金产品可供申购")
        else:
            # 展示当前母基金的持仓（持有的子基金份额）
            holdings = st.session_state.holdings.get(selected_mother, {})
            st.subheader(f"📦 {selected_mother} 当前持仓（子基金）")
            if holdings:
                holdings_df = pd.DataFrame([
                    {"子基金": name, "持有份额 (万元)": amt / 10000} for name, amt in holdings.items()
                ])
                st.dataframe(holdings_df, use_container_width=True)
            else:
                st.write("暂无持仓")

            # 申购/赎回操作
            st.subheader("申购/赎回子基金")
            target_child = st.selectbox("选择子基金", children, key="child_select")
            # 获取子基金当前净值（简化：净值固定为1，也可以让用户输入）
            nav = st.number_input(f"{target_child} 单位净值 (元)", min_value=0.0, value=1.0, step=0.01, key="nav")
            # 交易类型
            action = st.radio("操作类型", ["申购", "赎回"], horizontal=True)
            # 金额或份额
            if action == "申购":
                amount = st.number_input("申购金额 (万元)", min_value=0.0, step=1.0, key="sub_amount")
                if st.button("执行申购"):
                    if amount <= 0:
                        st.error("金额必须大于0")
                    else:
                        amount_yuan = amount * 10000
                        # 获取母基金当前余额（最后一条记录的期末余额）
                        mother_df = sheets_data[selected_mother]
                        mother_balance = mother_df["期末余额"].iloc[-1] if not mother_df.empty else 0
                        if mother_balance < amount_yuan:
                            st.error(f"母基金余额不足（当前余额 {mother_balance:,.0f} 元）")
                        else:
                            # 计算申购份额 = 金额 / 净值
                            shares = amount_yuan / nav
                            # 1. 母基金：现金流出
                            new_mother_balance = mother_balance - amount_yuan
                            add_cashflow_record(
                                selected_mother,
                                date.today(),
                                f"申购子基金 {target_child}",
                                0, amount_yuan, new_mother_balance
                            )
                            # 2. 子基金：现金流入，并增加期末余额
                            child_df = sheets_data[target_child]
                            child_balance = child_df["期末余额"].iloc[-1] if not child_df.empty else 0
                            new_child_balance = child_balance + amount_yuan
                            add_cashflow_record(
                                target_child,
                                date.today(),
                                f"被母基金 {selected_mother} 申购",
                                amount_yuan, 0, new_child_balance
                            )
                            # 3. 更新母基金持仓
                            current_hold = st.session_state.holdings.get(selected_mother, {})
                            current_hold[target_child] = current_hold.get(target_child, 0) + shares
                            st.session_state.holdings[selected_mother] = current_hold
                            # 刷新缓存
                            st.cache_data.clear()
                            st.success(
                                f"申购成功！母基金流出 {amount_yuan:,.0f} 元，子基金流入 {amount_yuan:,.0f} 元，持有份额增加 {shares:,.2f}")
                            st.rerun()
            else:  # 赎回
                # 赎回需要基于持仓份额
                current_shares = holdings.get(target_child, 0)
                if current_shares == 0:
                    st.warning(f"母基金未持有 {target_child}，无法赎回")
                else:
                    # 可以赎回全部或部分，按份额计算赎回金额
                    shares_to_redeem = st.number_input("赎回份额 (份)", min_value=0.0, max_value=float(current_shares),
                                                       step=1000.0, key="redeem_shares")
                    if st.button("执行赎回"):
                        if shares_to_redeem <= 0:
                            st.error("份额必须大于0")
                        else:
                            redeem_amount = shares_to_redeem * nav
                            # 1. 母基金：现金流入
                            mother_df = sheets_data[selected_mother]
                            mother_balance = mother_df["期末余额"].iloc[-1] if not mother_df.empty else 0
                            new_mother_balance = mother_balance + redeem_amount
                            add_cashflow_record(
                                selected_mother,
                                date.today(),
                                f"赎回子基金 {target_child}",
                                redeem_amount, 0, new_mother_balance
                            )
                            # 2. 子基金：现金流出
                            child_df = sheets_data[target_child]
                            child_balance = child_df["期末余额"].iloc[-1] if not child_df.empty else 0
                            new_child_balance = child_balance - redeem_amount
                            add_cashflow_record(
                                target_child,
                                date.today(),
                                f"被母基金 {selected_mother} 赎回",
                                0, redeem_amount, new_child_balance
                            )
                            # 3. 更新持仓
                            current_hold = st.session_state.holdings.get(selected_mother, {})
                            current_hold[target_child] -= shares_to_redeem
                            if current_hold[target_child] <= 0:
                                del current_hold[target_child]
                            st.session_state.holdings[selected_mother] = current_hold
                            st.cache_data.clear()
                            st.success(
                                f"赎回成功！母基金流入 {redeem_amount:,.0f} 元，子基金流出 {redeem_amount:,.0f} 元")
                            st.rerun()

# ---------- 6. 全局仪表盘：所有产品最新余额及趋势 ----------
st.header("📊 流动性仪表盘")
if all_products:
    latest_data = []
    for prod in all_products:
        df = sheets_data[prod]
        if not df.empty:
            latest_balance = df["期末余额"].iloc[-1]
            if pd.notna(latest_balance):
                latest_data.append({"产品": prod, "最新余额 (万元)": latest_balance / 10000})
    if latest_data:
        latest_df = pd.DataFrame(latest_data)
        col1, col2 = st.columns(2)
        with col1:
            st.metric("平均余额 (万元)", f"{latest_df['最新余额 (万元)'].mean():.2f}")
        with col2:
            st.metric("总余额 (万元)", f"{latest_df['最新余额 (万元)'].sum():.2f}")
        st.bar_chart(latest_df.set_index("产品"))

    # 趋势图（前5个产品）
    st.subheader("余额趋势（前5个产品）")
    plot_products = all_products[:5]
    trend_dfs = []
    for prod in plot_products:
        df = sheets_data[prod]
        if not df.empty:
            trend = df[["日期", "期末余额"]].copy()
            trend["产品"] = prod
            trend_dfs.append(trend)
    if trend_dfs:
        trend_all = pd.concat(trend_dfs, ignore_index=True)
        fig = px.line(trend_all, x="日期", y="期末余额", color="产品", title="产品期末余额趋势")
        st.plotly_chart(fig, use_container_width=True)

st.caption("提示：所有现金流修改都会自动保存到 Excel 的前5列。母子基金申购/赎回会生成新的交易记录并更新双方余额。")