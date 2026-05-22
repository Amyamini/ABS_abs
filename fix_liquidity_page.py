"""
恢复 page_product_liquidity.py 为独立运行脚本
移除 render_product_liquidity() 函数封装
"""

# 读取当前文件
with open('page_product_liquidity.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 处理文件内容
output_lines = []
in_function = False
function_removed = False

for i, line in enumerate(lines):
    # 跳过函数定义行
    if 'def render_product_liquidity():' in line:
        in_function = True
        function_removed = True
        continue
    
    # 跳过函数的文档字符串
    if in_function and '"""渲染产品流动性管理页面"""' in line:
        continue
    
    # 如果在函数内，移除一级缩进（4个空格）
    if in_function and line.startswith('    '):
        output_lines.append(line[4:])  # 移除前4个空格
    elif in_function and line.strip() == '':
        # 空行保持
        output_lines.append(line)
    else:
        output_lines.append(line)

# 写入修复后的文件
with open('page_product_liquidity.py', 'w', encoding='utf-8') as f:
    f.writelines(output_lines)

print("✅ page_product_liquidity.py 已恢复为独立运行脚本")
print(f"   - 移除了 render_product_liquidity() 函数封装")
print(f"   - 共处理 {len(lines)} 行代码")
