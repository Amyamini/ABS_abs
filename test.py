import ollama

# ====================== 【只改这里！】你的代码要求 ======================
CODE_REQUIREMENT = """
1. 编程语言：Python
2. 功能：寻找并优化同一目录下的page_product_liquidity.py
3. 格式要求：函数化编程，带类型注解，每行关键代码加中文注释
4. 附加要求：包含测试代码，代码简洁无冗余
"""
# =====================================================================

# 调用DeepSeek-Coder生成代码
response = ollama.chat(
    model="deepseek-coder:6.7b",  # 你下载的模型名
    messages=[
        # 系统提示：固定告诉模型你是专业代码生成器
        {"role": "system", "content": "你是专业的代码生成助手，严格按照用户要求生成干净、可直接运行、带完整注释的代码。"},
        {"role": "user", "content": CODE_REQUIREMENT}
    ],
    # 代码生成最优参数（不用改）
    options={
        "temperature": 0.2,    # 越低代码越稳定、越规范
        "num_predict": 2048    # 最大代码长度
    }
)

# 输出生成的代码
print("="*50 + "生成的代码" + "="*50)
print(response["message"]["content"])