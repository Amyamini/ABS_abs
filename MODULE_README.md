# ABS投资跟进系统 - 模块化重构说明

## 项目结构

本次重构将原来的单体 `app.py` 文件（1085行）拆分为多个模块，提高代码的可维护性和可读性。

### 文件列表

```
YKABS/
├── app.py                      # 主应用文件（187行）- 精简版
├── data_loader.py              # 数据加载模块（98行）
├── statistics.py               # 统计分析模块（127行）
├── page_core_metrics.py        # 核心投资指标页面（300行）
├── page_product_analysis.py    # 产品分析页面（163行）
├── page_asset_analysis.py      # 资产分析页面（150行）
└── page_trade_records.py       # 交易记录页面（98行）
```

### 模块说明

#### 1. **app.py** - 主应用入口
- 页面配置和CSS样式
- 侧边栏菜单
- 数据加载和预处理
- 路由控制（根据菜单选择显示不同页面）

#### 2. **data_loader.py** - 数据加载模块
- `load_data()`: 加载主数据文件 data.xlsx
- `load_trade()`: 加载交易记录文件
- `load_projects()`: 加载项目库文件
- `merge_data_with_projects()`: 关联主数据和项目库

#### 3. **statistics.py** - 统计分析模块
- `calculate_statistics()`: 计算核心统计指标
- `calculate_trade_records()`: 计算交易记录统计

#### 4. **page_core_metrics.py** - 核心投资指标页面
- 展示4个核心指标卡片
- 资产类型分布饼图
- 月度交易趋势图
- 规模余额图表
- 产品存续规模Top20

#### 5. **page_product_analysis.py** - 产品分析页面
- 产品持仓及收益明细表
- 支持产品和证券模糊搜索
- XIRR收益率计算
- 汇总统计

#### 6. **page_asset_analysis.py** - 资产分析页面
- 资产清单展示
- 多字段模糊筛选（项目名称、证券名称、资产类型等）
- 数据格式化（发行规模单位转换、日期格式化）
- 底部统计信息

#### 7. **page_trade_records.py** - 交易记录页面
- 交易记录明细表
- 时间、证券、买卖方筛选
- 交易份额和金额汇总

## 优势

1. **代码组织清晰**: 每个模块职责单一，易于理解和维护
2. **便于扩展**: 新增功能只需添加新模块或修改对应模块
3. **减少耦合**: 各模块之间通过函数调用交互，降低依赖
4. **提高可测试性**: 每个模块可以独立测试
5. **主文件精简**: app.py从1085行减少到187行，逻辑更清晰

## 运行方式

```bash
streamlit run app.py
```

## 注意事项

- 所有模块文件必须与 app.py 在同一目录下
- 确保已安装所有依赖包（streamlit, pandas, plotly, scipy等）
- 数据文件（data.xlsx, 交易记录.xlsx, 项目库.xlsx）需放在项目根目录
