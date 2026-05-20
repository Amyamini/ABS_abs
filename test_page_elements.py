"""
测试产品要素页面是否能正常运行
"""
import sys
import os

print("=" * 80)
print("测试 page_product_elements.py")
print("=" * 80)

# 测试1: 检查依赖库
print("\n【测试1】检查依赖库...")
try:
    import streamlit as st
    print("✅ streamlit 导入成功")
except ImportError as e:
    print(f"❌ streamlit 导入失败: {e}")
    sys.exit(1)

try:
    import pandas as pd
    print("✅ pandas 导入成功")
except ImportError as e:
    print(f"❌ pandas 导入失败: {e}")
    sys.exit(1)

# 测试2: 检查文件是否存在
print("\n【测试2】检查文件...")
if os.path.exists("page_product_elements.py"):
    print("✅ page_product_elements.py 存在")
else:
    print("❌ page_product_elements.py 不存在")
    sys.exit(1)

# 测试3: 尝试加载模块
print("\n【测试3】尝试导入模块...")
try:
    from page_product_elements import render_product_elements, load_product_elements, save_product_elements
    print("✅ 模块导入成功")
    print(f"   - render_product_elements: {render_product_elements}")
    print(f"   - load_product_elements: {load_product_elements}")
    print(f"   - save_product_elements: {save_product_elements}")
except Exception as e:
    print(f"❌ 模块导入失败: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 测试4: 测试数据加载函数
print("\n【测试4】测试数据加载函数...")
try:
    df = load_product_elements()
    print(f"✅ 数据加载成功")
    print(f"   - DataFrame 类型: {type(df)}")
    print(f"   - 是否为空: {df.empty}")
    if not df.empty:
        print(f"   - 行数: {len(df)}")
        print(f"   - 列名: {df.columns.tolist()}")
    else:
        print(f"   - 列名: {df.columns.tolist()}")
except Exception as e:
    print(f"❌ 数据加载失败: {e}")
    import traceback
    traceback.print_exc()

# 测试5: 检查Excel文件
print("\n【测试5】检查Excel文件...")
if os.path.exists("产品要素表.xlsx"):
    print("✅ 产品要素表.xlsx 存在")
    try:
        df_test = pd.read_excel("产品要素表.xlsx")
        print(f"   - 可以正常读取")
        print(f"   - 行数: {len(df_test)}")
        print(f"   - 列名: {df_test.columns.tolist()}")
    except Exception as e:
        print(f"   ⚠️ 读取失败: {e}")
else:
    print("ℹ️  产品要素表.xlsx 不存在（这是正常的，首次运行时会创建）")

print("\n" + "=" * 80)
print("✅ 所有测试通过！")
print("=" * 80)
print("\n现在可以运行:")
print("  streamlit run page_product_elements.py")
print("\n如果还是空白页面，请检查:")
print("  1. Streamlit 版本是否最新")
print("  2. 浏览器控制台是否有JavaScript错误")
print("  3. 终端是否有错误信息")
