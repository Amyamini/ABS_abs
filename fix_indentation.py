# -*- coding: utf-8 -*-
"""
修复 page_product_liquidity.py 的缩进问题
移除第11行之后的所有行的前4个空格缩进
"""

input_file = 'page_product_liquidity.py'
output_file = 'page_product_liquidity_fixed.py'

with open(input_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

fixed_lines = []
for i, line in enumerate(lines):
    # 前11行（索引0-10）保持不变
    if i < 11:
        fixed_lines.append(line)
    else:
        # 从第12行开始，如果以4个空格开头，则移除
        if line.startswith('    '):
            fixed_lines.append(line[4:])
        else:
            fixed_lines.append(line)

# 写入修复后的文件
with open(output_file, 'w', encoding='utf-8') as f:
    f.writelines(fixed_lines)

print(f"✅ 修复完成！")
print(f"   - 原始文件: {input_file}")
print(f"   - 修复文件: {output_file}")
print(f"   - 总行数: {len(lines)}")
print(f"\n请检查 {output_file} 是否正确，然后替换原文件。")
