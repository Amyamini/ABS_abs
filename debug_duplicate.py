"""
调试脚本：检查为什么会有重复行
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

# 取最新日期
df_latest_date = df_copy.groupby(["产品名称", "证券名称"])["日期"].max().reset_index()
latest_date = df_latest_date['日期'].values[0]
print(f"\n最新日期: {latest_date}")

# 取最新日期的所有记录
df_latest = df_copy.merge(df_latest_date, on=["产品名称", "证券名称", "日期"], how="inner")

print(f"\n最新日期的记录数: {len(df_latest)}")
print(f"\n最新日期的所有字段:")
print(df_latest.columns.tolist())

print(f"\n【最新日期的完整数据】")
for idx, row in df_latest.iterrows():
    print(f"\n记录 {idx + 1}:")
    for col in df_latest.columns:
        print(f"  {col}: {row[col]}")

print("\n" + "=" * 100)
print("分析：如果最新日期有多条记录，可能是因为：")
print("1. 同一天有多笔交易")
print("2. 数据中有其他维度字段（如不同账户、不同批次等）")
print("=" * 100)
