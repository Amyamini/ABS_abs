# 产品参数和流动性管理表页面分离说明

## 功能概述

已将产品参数和流动性管理表分别展示在两个独立的页面中，方便用户分别管理和查看。

## 页面结构

### 1. 产品参数页面
- **菜单名称**: 产品参数
- **文件**: `page_product_elements.py`
- **功能**:
  - 管理产品要素信息（申购赎回到账时间等）
  - 支持添加、编辑、删除产品信息
  - 提供快速查询工具
  - 数据保存在 `产品要素表.xlsx`

### 2. 流动性管理表页面
- **菜单名称**: 流动性管理表
- **文件**: `page_product_liquidity.py`
- **功能**:
  - 管理产品现金流明细
  - 母子基金联动操作
  - 流动性仪表盘展示
  - 数据保存在 `基金流动性管理总表.xlsx`

## 使用方法

### 启动应用
```bash
streamlit run app.py
```

### 访问页面
1. 在侧边栏菜单中选择 "产品参数" 或 "流动性管理表"
2. 根据需要执行相应操作

## 技术实现

### 主要修改

1. **app.py**
   - 添加了两个新菜单项："产品参数" 和 "流动性管理表"
   - 导入了 `render_product_elements` 和 `render_product_liquidity` 函数
   - 添加了相应的路由处理逻辑

2. **page_product_liquidity.py**
   - 将原有代码封装到 `render_product_liquidity()` 函数中
   - 保持所有原有功能不变
   - 支持作为模块被主应用调用

3. **page_product_elements.py**
   - 已有 `render_product_elements()` 函数，无需修改
   - 可直接被主应用调用

## 测试验证

运行测试脚本验证功能：
```bash
python test_page_separation.py
```

## 注意事项

1. 两个页面使用不同的数据文件，互不影响
2. 产品参数页面主要用于配置产品基本信息
3. 流动性管理表页面用于日常现金流管理和监控
4. 所有数据修改都会自动保存到对应的Excel文件中

## 文件清单

- `app.py` - 主应用文件（已修改）
- `page_product_elements.py` - 产品参数页面（无需修改）
- `page_product_liquidity.py` - 流动性管理表页面（已修改为函数形式）
- `test_page_separation.py` - 测试脚本（新增）
- `PAGE_SEPARATION_README.md` - 本说明文档（新增）
