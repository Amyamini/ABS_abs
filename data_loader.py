"""
数据加载与预处理模块
负责加载和清洗所有数据源
"""
import pandas as pd
import streamlit as st


# 1. 加载主数据 data.xlsx
@st.cache_data
def load_data():
    """加载并预处理主数据文件"""
    try:
        df = pd.read_excel("data.xlsx")

        # 日期标准化
        if "日期" in df.columns:
            df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
            df = df.dropna(subset=["日期"])
            df["年月"] = df["日期"].dt.to_period("M").astype(str)

        # 本金-现金流 字段标准化（后面要算剩余本金）
        if "本金-现金流" in df.columns:
            df["本金-现金流"] = pd.to_numeric(df["本金-现金流"], errors="coerce").fillna(0)

        return df
    except Exception as e:
        st.error(f"主数据加载失败：{str(e)}")
        st.stop()


# 2. 加载交易记录
@st.cache_data
def load_trade():
    """加载并预处理交易记录文件"""
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

        # 4. 产品1 字段预处理（后面要画空值柱状图）
        if "产品1" in trades.columns:
            trades["产品1_为空"] = trades["产品1"].isnull()

        return trades
    except Exception as e:
        st.error(f"交易记录加载失败：{str(e)}")
        st.stop()


# 3. 加载项目库
@st.cache_data
def load_projects():
    """加载项目库文件"""
    try:
        projects = pd.read_excel("项目库.xlsx")
        return projects
    except Exception as e:
        st.error(f"项目库加载失败：{str(e)}")
        st.stop()


def merge_data_with_projects(df, df_projects):
    """
    关联data表和项目库
    
    Args:
        df: 主数据DataFrame
        df_projects: 项目库DataFrame
        
    Returns:
        合并后的DataFrame
    """
    # 通过"证券名称"左关联（保留data所有数据，关联项目库的补充信息）
    df_merged = pd.merge(
        df,  # 主表：data
        df_projects,  # 关联表：项目库
        on="证券名称",  # 关联字段：证券名称
        how="left",  # 左连接：保留data所有数据，项目库无匹配则为空
        suffixes=("", "_项目库")  # 重复字段加后缀区分
    )
    
    return df_merged
