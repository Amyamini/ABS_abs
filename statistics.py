"""
统计分析模块
负责计算各种统计指标和分析数据
"""
import pandas as pd


def calculate_statistics(df):
    """
    计算核心统计指标
    
    Args:
        df: 过滤后的DataFrame
        
    Returns:
        包含各项统计指标的字典
    """
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


def calculate_trade_records(df):
    """
    计算交易记录统计指标
    
    Args:
        df: 交易记录DataFrame
        
    Returns:
        包含交易记录统计的字典
    """
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
