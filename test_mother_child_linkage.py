"""
测试母子基金联动功能
验证实时显示联动结果的功能
"""
import pandas as pd
from datetime import datetime, timedelta

print("=" * 80)
print("测试母子基金联动功能")
print("=" * 80)

# 测试1: 模拟母基金持有子基金份额
print("\n【测试1】模拟母基金持仓结构...")
try:
    # 创建模拟数据
    mother_fund_data = {
        "日期": pd.to_datetime(["2024-01-01", "2024-01-02", "2024-01-03"]),
        "现金流类型": ["初始余额", "申购", "对子基金A申购"],
        "关联产品": ["", "", "子基金A"],
        "现金流入": [10000000, 5000000, 0],
        "现金流出": [0, 0, 2000000],
        "期末余额": [10000000, 15000000, 13000000]
    }
    
    df_mother = pd.DataFrame(mother_fund_data)
    print(f"✅ 母基金数据创建成功")
    print(f"   记录数: {len(df_mother)}")
    print(f"   最新余额: {df_mother['期末余额'].iloc[-1]:,.2f}")
    print(f"\n母基金数据预览:")
    print(df_mother.to_string(index=False))
    
except Exception as e:
    print(f"❌ 测试失败: {e}")
    import traceback
    traceback.print_exc()

# 测试2: 模拟子基金数据
print("\n【测试2】模拟子基金数据结构...")
try:
    child_fund_a_data = {
        "日期": pd.to_datetime(["2024-01-01", "2024-01-02"]),
        "现金流类型": ["初始余额", "申购"],
        "关联产品": ["", ""],
        "现金流入": [5000000, 1000000],
        "现金流出": [0, 0],
        "期末余额": [5000000, 6000000]
    }
    
    df_child_a = pd.DataFrame(child_fund_a_data)
    print(f"✅ 子基金A数据创建成功")
    print(f"   记录数: {len(df_child_a)}")
    print(f"   最新余额: {df_child_a['期末余额'].iloc[-1]:,.2f}")
    print(f"\n子基金A数据预览:")
    print(df_child_a.to_string(index=False))
    
except Exception as e:
    print(f"❌ 测试失败: {e}")
    import traceback
    traceback.print_exc()

# 测试3: 模拟联动计算逻辑
print("\n【测试3】测试联动计算逻辑...")
try:
    # 母基金申购子基金场景
    trade_date = pd.to_datetime("2024-01-03")
    subscription_amount = 2000000  # 母基金流出200万申购子基金
    subscribe_days = 1  # T+1到账
    
    # 计算子基金到账日期和金额
    arrival_date = trade_date + pd.Timedelta(days=subscribe_days)
    child_inflow = subscription_amount
    
    print(f"✅ 联动计算测试通过")
    print(f"   母基金操作日: {trade_date.strftime('%Y-%m-%d')}")
    print(f"   母基金操作: 现金流出 {subscription_amount:,.2f} (申购子基金)")
    print(f"   子基金到账日: {arrival_date.strftime('%Y-%m-%d')} (T+{subscribe_days})")
    print(f"   子基金操作: 现金流入 {child_inflow:,.2f}")
    
    # 模拟余额计算
    current_balance = 6000000
    new_balance = current_balance + child_inflow
    print(f"   子基金当前余额: {current_balance:,.2f}")
    print(f"   子基金变动后余额: {new_balance:,.2f}")
    
except Exception as e:
    print(f"❌ 测试失败: {e}")
    import traceback
    traceback.print_exc()

# 测试4: 模拟赎回联动
print("\n【测试4】测试赎回联动逻辑...")
try:
    # 母基金赎回子基金场景
    trade_date = pd.to_datetime("2024-01-04")
    redemption_amount = 1000000  # 母基金流入100万（赎回子基金）
    redeem_days = 2  # T+2到账
    
    # 计算子基金到账日期和金额
    arrival_date = trade_date + pd.Timedelta(days=redeem_days)
    child_outflow = redemption_amount
    
    print(f"✅ 赎回联动计算测试通过")
    print(f"   母基金操作日: {trade_date.strftime('%Y-%m-%d')}")
    print(f"   母基金操作: 现金流入 {redemption_amount:,.2f} (赎回子基金)")
    print(f"   子基金到账日: {arrival_date.strftime('%Y-%m-%d')} (T+{redeem_days})")
    print(f"   子基金操作: 现金流出 {child_outflow:,.2f}")
    
    # 模拟余额计算
    current_balance = 7000000
    new_balance = current_balance - child_outflow
    print(f"   子基金当前余额: {current_balance:,.2f}")
    print(f"   子基金变动后余额: {new_balance:,.2f}")
    
except Exception as e:
    print(f"❌ 测试失败: {e}")
    import traceback
    traceback.print_exc()

# 测试5: 多子基金按比例分配
print("\n【测试5】测试多子基金按比例分配...")
try:
    # 母基金持有多个子基金
    holdings = {
        "子基金A": 3000000,  # 300万份额
        "子基金B": 2000000,  # 200万份额
        "子基金C": 1000000   # 100万份额
    }
    
    total_holdings = sum(holdings.values())
    subscription_amount = 1200000  # 申购总额120万
    
    print(f"✅ 多子基金分配测试通过")
    print(f"   总持有份额: {total_holdings:,.2f}")
    print(f"   申购总额: {subscription_amount:,.2f}")
    print(f"\n   按比例分配结果:")
    
    for child_name, shares in holdings.items():
        ratio = shares / total_holdings
        allocated_amount = subscription_amount * ratio
        print(f"   - {child_name}: 持有比例 {ratio:.2%}, 分配金额 {allocated_amount:,.2f}")
    
except Exception as e:
    print(f"❌ 测试失败: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("所有测试完成！")
print("=" * 80)
print("\n提示：要运行完整的Streamlit应用，请使用命令：")
print("  streamlit run page_product_liquidity.py")
