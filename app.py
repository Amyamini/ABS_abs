from datetime import datetime
import streamlit as st

# 导入自定义模块
from data_loader import load_data, load_trade, load_projects, merge_data_with_projects
from statistics import calculate_statistics, calculate_trade_records
from page_core_metrics import render_core_metrics
from page_product_analysis import render_product_analysis
from page_asset_analysis import render_asset_analysis
from page_trade_records import render_trade_records
from page_product_elements import render_product_elements


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
# 数据加载与预处理
# ----------------------

# 1. 加载主数据 data.xlsx
df = load_data()

# 2. 加载交易记录
trades = load_trade()

# 3. 加载项目库
df_projects = load_projects()

# =========== 关联data和项目库==============
df_merged = merge_data_with_projects(df, df_projects)


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
    
    menu = st.radio(
        label="",  # 隐藏标题
        options=[
            "核心投资指标",
            "产品分析",
            "资产分析",
            "交易记录",
            "产品参数",
        ],
        label_visibility="collapsed"
    )
    
    st.markdown("---")

# ----------------------
# 数据过滤
# ----------------------
# 所有页面使用相同的数据
df_filtered = df_merged.copy()
trades_filtered = trades.copy()
df_projects_filtered = df_projects.copy()

# ----------------------
# 核心统计计算
# ----------------------
stats = calculate_statistics(df_filtered)
trade_records = calculate_trade_records(trades_filtered)

# ----------------------
# 主页面标题
# ----------------------
st.markdown("<h1 class='main-header'>ABS投资跟进</h1>", unsafe_allow_html=True)
st.markdown(f"""
<div style='text-align: right; font-size: 14px; margin-bottom: 10px; font-weight: normal;'>
    数据时间范围：{stats['date_range']}
</div>
""", unsafe_allow_html=True)
# 超紧凑分割线（间距极小）
st.markdown("<hr style='margin: 5px 0; border: none; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# ----------------------
# 核心投资指标页面
# ----------------------
if menu == "核心投资指标":
    render_core_metrics(stats, df_filtered, trades_filtered)

# ----------------------
# 产品分析页面
# ----------------------
elif menu == "产品分析":
    render_product_analysis(df_filtered)

# ----------------------
# 资产分析页面
# ----------------------
elif menu == "资产分析":
    render_asset_analysis(df_projects_filtered)

# ----------------------
# 交易记录页面
# ----------------------
elif menu == "交易记录":
    render_trade_records(trades_filtered)

# ----------------------
# 产品参数页面
# ----------------------
elif menu == "产品参数":
    render_product_elements()


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
