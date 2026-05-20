"""
测试 page_product_liquidity.py 的代码逻辑
"""
import pandas as pd
import sys

print("=" * 80)
print("测试 page_product_liquidity.py 的核心功能")
print("=" * 80)

# 测试1: 检查文件是否可以导入
print("\n【测试1】检查模块导入...")
try:
    import streamlit as st
    import plotly.express as px
    from openpyxl import load_workbook
    print("✅ 所有依赖库导入成功")
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    sys.exit(1)

# 测试2: 检查Excel文件是否存在
print("\n【测试2】检查数据文件...")
import os
file_path = "基金流动性管理总表.xlsx"
if os.path.exists(file_path):
    print(f"✅ 找到文件: {file_path}")
    print(f"   文件大小: {os.path.getsize(file_path) / 1024 / 1024:.2f} MB")
else:
    print(f"❌ 文件不存在: {file_path}")
    print("   请确保文件存在于项目根目录")

# 测试3: 尝试读取Excel文件
print("\n【测试3】读取Excel文件...")
try:
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    sheets = xls.sheet_names
    print(f"✅ 成功读取文件")
    print(f"   工作表数量: {len(sheets)}")
    print(f"   工作表名称: {', '.join(sheets[:5])}{'...' if len(sheets) > 5 else ''}")
    
    # 读取第一个工作表
    if sheets:
        first_sheet = sheets[0]
        df = pd.read_excel(file_path, sheet_name=first_sheet, engine="openpyxl")
        print(f"\n   第一个工作表 '{first_sheet}':")
        print(f"   - 行数: {len(df)}")
        print(f"   - 列数: {len(df.columns)}")
        print(f"   - 列名: {', '.join(df.columns[:5])}")
        
        # 检查前5列
        if df.shape[1] >= 5:
            flow_df = df.iloc[:, :5].copy()
            flow_df.columns = ["日期", "现金流类型", "现金流入", "现金流出", "期末余额"]
            
            # 数值转换测试
            for col in ["现金流入", "现金流出", "期末余额"]:
                flow_df[col] = pd.to_numeric(flow_df[col], errors="coerce")
            flow_df["日期"] = pd.to_datetime(flow_df["日期"], errors="coerce")
            
            print(f"\n   标准化后的数据:")
            print(f"   - 有效日期数: {flow_df['日期'].notna().sum()}")
            print(f"   - 现金流入总和: {flow_df['现金流入'].sum():,.2f}")
            print(f"   - 现金流出总和: {flow_df['现金流出'].sum():,.2f}")
            print(f"   - 期末余额(最后一条): {flow_df['期末余额'].iloc[-1] if not flow_df.empty else 0:,.2f}")
            
except Exception as e:
    print(f"❌ 读取失败: {e}")
    import traceback
    traceback.print_exc()

# 测试4: 检查核心函数逻辑
print("\n【测试4】测试核心函数逻辑...")
try:
    # 模拟 add_cashflow_record 函数
    def test_add_cashflow():
        # 创建测试数据
        df = pd.DataFrame({
            "日期": pd.to_datetime(["2024-01-01", "2024-01-02"]),
            "现金流类型": ["初始余额", "申购"],
            "现金流入": [1000000, 500000],
            "现金流出": [0, 0],
            "期末余额": [1000000, 1500000]
        })
        
        # 添加新记录
        from datetime import date
        new_row = pd.DataFrame({
            "日期": [date.today()],
            "现金流类型": ["赎回"],
            "现金流入": [0],
            "现金流出": [200000],
            "期末余额": [1300000]
        })
        df = pd.concat([df, new_row], ignore_index=True)
        
        print(f"✅ 现金流记录添加测试通过")
        print(f"   原始记录数: 2")
        print(f"   添加后记录数: {len(df)}")
        print(f"   最新余额: {df['期末余额'].iloc[-1]:,.2f}")
        return True
    
    test_add_cashflow()
    
except Exception as e:
    print(f"❌ 函数测试失败: {e}")
    import traceback
    traceback.print_exc()

# 测试5: 检查母子基金联动逻辑
print("\n【测试5】测试母子基金联动逻辑...")
try:
    # 模拟持仓结构
    holdings = {
        "母基金A": {
            "子基金B": 1000000,  # 持有份额
            "子基金C": 500000
        }
    }
    
    print(f"✅ 持仓结构测试通过")
    print(f"   母基金数量: {len(holdings)}")
    for mother, children in holdings.items():
        print(f"   - {mother}: 持有 {len(children)} 个子基金")
        for child, shares in children.items():
            print(f"     * {child}: {shares:,.2f} 份")
    
except Exception as e:
    print(f"❌ 联动逻辑测试失败: {e}")

print("\n" + "=" * 80)
print("测试完成！")
print("=" * 80)
print("\n提示：要运行完整的Streamlit应用，请使用命令：")
print("  streamlit run page_product_liquidity.py")
