# -*- coding: utf-8 -*-
"""
一键修复 page_product_liquidity.py 的缩进问题
"""

# 读取文件
with open('page_product_liquidity.py', 'r', encoding='utf-8') as f:
    content = f.read()
    lines = content.split('\n')

# 修复缩进：从第12行（索引11）开始，移除每行开头的4个空格
fixed_lines = []
for i, line in enumerate(lines):
    if i >= 11 and line.startswith('    '):
        # 移除前4个空格
        fixed_lines.append(line[4:])
    else:
        fixed_lines.append(line)

# 写回文件
with open('page_product_liquidity.py', 'w', encoding='utf-8') as f:
    f.write('\n'.join(fixed_lines))

print("✅ 修复完成！page_product_liquidity.py 已恢复为独立运行脚本")
print(f"   - 共处理 {len(lines)} 行代码")
print(f"   - 移除了第12行之后的多余缩进")
