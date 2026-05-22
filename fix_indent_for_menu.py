# -*- coding: utf-8 -*-
"""
将 page_product_liquidity.py 中第239行之后的所有内容缩进4个空格
使其成为 elif 分支的一部分
"""

input_file = 'page_product_liquidity.py'

with open(input_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 从第239行（索引238）开始，为每行添加4个空格缩进
output_lines = []
for i, line in enumerate(lines):
    if i >= 238 and line.strip():  # 非空行才添加缩进
        output_lines.append('    ' + line)
    else:
        output_lines.append(line)

with open(input_file, 'w', encoding='utf-8') as f:
    f.writelines(output_lines)

print(f"✅ 修复完成！共处理 {len(lines)} 行代码")
print(f"   - 从第239行开始添加了4空格缩进")
