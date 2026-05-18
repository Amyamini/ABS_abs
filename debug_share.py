"""
调试脚本：验证元康盛景28号-惠民3次的持仓份额计算
"""
import pandas as pd

# 加载数据
df = pd.read_excel("data.xlsx")

# 筛选目标数据
target_product = "元康盛景28号"
target_security = "惠民3次"

df_target = df[(df["产品名称"] == target_product) & (df["证券名称"] == target_security)].copy()

print("=" * 80)
print(f"产品: {target_product}")
print(f"证券: {target_security}")
print("=" * 80)

# 检查是否有交易份额字段
if "交易份额（万份）" in df_target.columns:
    print("\n【原始数据 - 交易份额明细】")
    print(df_target[["日期", "产品名称", "证券名称", "交易份额（万份）", "本金-现金流"]].to_string())
    
    print(f"\n交易份额总和: {df_target['交易份额（万份）'].sum()}")
    print(f"交易笔数: {len(df_target)}")
    
    # 按日期排序
    df_sorted = df_target.sort_values("日期")
    print("\n【按日期排序后的数据】")
    print(df_sorted[["日期", "交易份额（万份）", "本金-现金流"]].to_string())
else:
    print("\n⚠️ 数据中没有'交易份额（万份）'字段")
    print(f"可用字段: {df_target.columns.tolist()}")

# 检查持仓份额字段
if "持仓份额（万份）" in df_target.columns:
    print("\n【持仓份额字段数据】")
    print(df_target[["日期", "持仓份额（万份）"]].to_string())
    
    # 最新日期的持仓份额
    latest_date = df_target["日期"].max()
    latest_hold = df_target[df_target["日期"] == latest_date]["持仓份额（万份）"].values
    print(f"\n最新日期: {latest_date}")
    print(f"最新持仓份额: {latest_hold}")

print("\n" + "=" * 80)
