"""
调试脚本：模拟产品分析页面的计算流程
"""
import pandas as pd

# 加载数据
df = pd.read_excel("data.xlsx")

# 日期标准化
if "日期" in df.columns:
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"])

# 筛选目标数据
target_product = "元康盛景28号"
target_security = "惠民3次"

df_copy = df[(df["产品名称"] == target_product) & (df["证券名称"] == target_security)].copy()

print("=" * 100)
print(f"产品: {target_product}")
print(f"证券: {target_security}")
print("=" * 100)

if "交易份额（万份）" not in df_copy.columns:
    print("\n⚠️ 数据中没有'交易份额（万份）'字段")
    print(f"可用字段: {df_copy.columns.tolist()}")
else:
    print(f"\n【步骤0】原始数据 ({len(df_copy)} 条记录)")
    print(df_copy[["日期", "交易份额（万份）", "本金-现金流"]].to_string())
    
    # ======================================================================
    # 步骤1：取最新时间点数据
    # ======================================================================
    df_latest_date = df_copy.groupby(["产品名称", "证券名称"])["日期"].max().reset_index()
    print(f"\n【步骤1】最新日期: {df_latest_date['日期'].values[0]}")
    
    df_latest = df_copy.merge(df_latest_date, on=["产品名称", "证券名称", "日期"], how="inner")
    print(f"最新日期的记录数: {len(df_latest)}")
    if len(df_latest) > 0:
        print(df_latest[["日期", "交易份额（万份）", "本金-现金流"]].to_string())
    
    # ======================================================================
    # 步骤2：持仓份额 = 交易份额求和
    # ======================================================================
    share_sum = df_copy.groupby(["产品名称", "证券名称"])["交易份额（万份）"].sum().reset_index()
    print(f"\n【步骤2】所有历史数据的交易份额总和: {share_sum['交易份额（万份）'].values[0]}")
    
    share_sum.rename(columns={"交易份额（万份）": "持仓份额"}, inplace=True)
    df_latest = df_latest.merge(share_sum, on=["产品名称", "证券名称"], how="left")
    
    print(f"Merge后的df_latest记录数: {len(df_latest)}")
    if len(df_latest) > 0:
        print(df_latest[["日期", "持仓份额"]].to_string())
    
    # ======================================================================
    # 步骤3：再次分组聚合（问题所在！）
    # ======================================================================
    print(f"\n【步骤3】再次groupby前的持仓份额: {df_latest['持仓份额'].values}")
    
    df_final = df_latest.groupby(["产品名称", "证券名称"], as_index=False).agg({
        "持仓份额": "sum",
    })
    
    print(f"再次groupby后的持仓份额: {df_final['持仓份额'].values[0]}")
    
    print("\n" + "=" * 100)
    print(f"❓ 为什么不是540？")
    print(f"   - 第67行计算的总和: {share_sum['持仓份额'].values[0]}")
    print(f"   - 最新日期的记录数: {len(df_latest)}")
    print(f"   - 第115行再次求和后的结果: {df_final['持仓份额'].values[0]}")
    print(f"   - 如果最新日期有{len(df_latest)}条记录，每条记录的持仓份额都是{share_sum['持仓份额'].values[0]}")
    print(f"   - 再次sum后会变成: {share_sum['持仓份额'].values[0]} × {len(df_latest)} = {share_sum['持仓份额'].values[0] * len(df_latest)}")
    print("=" * 100)
