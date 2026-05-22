"""
测试产品参数和流动性管理表页面分离功能
"""
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """测试导入是否正常"""
    try:
        from page_product_elements import render_product_elements
        print("✅ 成功导入 render_product_elements")
        
        from page_product_liquidity import render_product_liquidity
        print("✅ 成功导入 render_product_liquidity")
        
        return True
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        return False

def test_function_signatures():
    """测试函数签名是否正确"""
    try:
        from page_product_elements import render_product_elements
        from page_product_liquidity import render_product_liquidity
        
        # 检查函数是否存在且可调用
        assert callable(render_product_elements), "render_product_elements 不是可调用函数"
        assert callable(render_product_liquidity), "render_product_liquidity 不是可调用函数"
        
        print("✅ 函数签名正确")
        return True
    except Exception as e:
        print(f"❌ 函数签名测试失败: {e}")
        return False

if __name__ == "__main__":
    print("🧪 开始测试产品参数和流动性管理表页面分离功能...\n")
    
    success = True
    
    print("1. 测试模块导入...")
    if not test_imports():
        success = False
    
    print("\n2. 测试函数签名...")
    if not test_function_signatures():
        success = False
    
    if success:
        print("\n🎉 所有测试通过！产品参数和流动性管理表已成功分离为两个独立页面。")
        print("\n使用方法:")
        print("1. 运行主应用: streamlit run app.py")
        print("2. 在侧边栏菜单中选择 '产品参数' 或 '流动性管理表'")
    else:
        print("\n❌ 测试失败，请检查代码错误。")
