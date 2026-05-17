"""
资产分析页面模块
展示资产清单和项目信息
"""
import streamlit as st
import pandas as pd


def render_asset_analysis(df_projects):
    """
    渲染资产分析页面
    
    Args:
        df_projects: 项目库数据
    """
    # ====================== 带筛选的项目清单模块（优化版） ======================
    st.markdown("""
    <div class='card'>
        <div class='card-title' style='color: black;'>资产清单</div>
    """, unsafe_allow_html=True)

    # 1. 加载项目库Excel
    try:
        df_projects_copy = df_projects.copy()

        # 2. 多字段模糊筛选器（保持原有布局）
        st.markdown("<div style='margin-bottom:15px;'>", unsafe_allow_html=True)
        # 第一行：项目名称 + 证券名称
        filter_col1, filter_col2 = st.columns(2, gap="small")
        with filter_col1:
            search_project = st.text_input("项目名称", placeholder="输入项目关键词...")
        with filter_col2:
            search_security = st.text_input("证券名称", placeholder="输入证券关键词...")

        # 第二行：资产类型 + 主体所属 + 资产类型二
        filter_col3, filter_col4, filter_col5 = st.columns(3, gap="small")
        with filter_col3:
            search_asset1 = st.text_input("资产类型", placeholder="输入资产类型关键词...")
        with filter_col4:
            search_subject = st.text_input("主体所属", placeholder="输入主体关键词...")
        with filter_col5:
            search_asset2 = st.text_input("资产类型二", placeholder="输入资产类型二关键词...")
        st.markdown("</div>", unsafe_allow_html=True)

        # 3. 应用模糊筛选
        if search_project:
            df_projects_copy = df_projects_copy[
                df_projects_copy["项目名称"].str.contains(search_project, na=False, case=False)
            ]
        if search_security:
            df_projects_copy = df_projects_copy[
                df_projects_copy["证券名称"].str.contains(search_security, na=False, case=False)
            ]
        if search_asset1:
            df_projects_copy = df_projects_copy[
                df_projects_copy["资产类型"].str.contains(search_asset1, na=False, case=False)
            ]
        if search_subject:
            df_projects_copy = df_projects_copy[
                df_projects_copy["主体所属"].str.contains(search_subject, na=False, case=False)
            ]
        if search_asset2:
            df_projects_copy = df_projects_copy[
                df_projects_copy["资产类型二"].str.contains(search_asset2, na=False, case=False)
            ]

        # 4. 核心优化：数据格式化（发行规模单位转换+固定日期格式）
        if not df_projects_copy.empty:
            # ---------------------- 优化1：发行规模（元→万元）转换 ----------------------
            # 先确保是数值类型，非数值转为 NaN
            if "发行规模（元）" in df_projects_copy.columns:
                # 转换为数值（处理可能存在的千分位逗号）
                df_projects_copy["发行规模（元）_numeric"] = df_projects_copy["发行规模（元）"].apply(
                    lambda x: float(str(x).replace(',', '')) if pd.notna(x) and str(x).replace(',', '').replace('.', '').isdigit() else None
                )
                
                # 单位转换：元 → 万元
                df_projects_copy["发行规模（万元）"] = df_projects_copy["发行规模（元）_numeric"] / 10000

                # 格式化：保留2位小数 + 千分位分隔符，空值显示"—"
                df_projects_copy["发行规模（万元）_formatted"] = df_projects_copy["发行规模（万元）"].apply(
                    lambda x: f"{x:,.2f}" if pd.notna(x) else "—"
                )
                
                # 删除临时列和原列
                df_projects_copy = df_projects_copy.drop(columns=["发行规模（元）", "发行规模（元）_numeric", "发行规模（万元）"])
                # 重命名格式化列为最终显示列
                df_projects_copy = df_projects_copy.rename(columns={"发行规模（万元）_formatted": "发行规模（万元）"})

            # ---------------------- 优化2：日期格式强制为 yyyy-mm-dd ----------------------
            # 识别所有含"日期"的列（如发行日期、到期日期等）
            date_cols = [col for col in df_projects_copy.columns if "预期到期日" in col or "清算日期" in col]
            for col in date_cols:
                # 强制转换为datetime，无效日期显示空值，最终格式固定为yyyy-mm-dd
                df_projects_copy[col] = pd.to_datetime(
                    df_projects_copy[col],
                    errors="coerce"  # 无法转换的日期设为NaT
                ).dt.strftime("%Y-%m-%d")  # 固定格式输出
                # 空值（NaT）显示为"—"，避免显示"NaT"
                df_projects_copy[col] = df_projects_copy[col].replace("NaT", "—")

            # ---------------------- 原有：预期收益率百分比格式化 ----------------------
            if "预期收益率" in df_projects_copy.columns:
                df_projects_copy["预期收益率"] = df_projects_copy["预期收益率"].apply(
                    lambda x: f"{x * 100:.2f}%" if pd.notna(x) and isinstance(x, (int, float)) else "—"
                )

            # 5. 显示美化表格（与交易记录表格高度一致）
            st.dataframe(
                df_projects_copy.style
                .set_properties(**{"text-align": "left", "font-size": "13px", "white-space": "nowrap"})
                .set_table_styles([{"selector": "th", "props": [("text-align", "left"), ("font-weight", "bold")]}]),
                width="stretch",
                height=450,  # 与交易记录表格保持一致
                hide_index=True
            )

            # 6. 底部统计（新增发行规模汇总）
            total_filtered = len(df_projects_copy)
            stats_text = f"筛选后项目总数：<span style='font-weight:bold; color:#1e40af;'>{total_filtered} 个</span>"

            # 新增：发行规模（万元）汇总（仅统计有效数值）
            if "发行规模（万元）" in df_projects_copy.columns:
                # 提取有效数值（排除"—"），并去除千分位逗号后转换为float
                valid_scale = df_projects_copy["发行规模（万元）"][df_projects_copy["发行规模（万元）"] != "—"]
                if not valid_scale.empty:
                    # 去除千分位逗号并转换为float
                    total_scale = valid_scale.apply(lambda x: float(str(x).replace(',', ''))).sum()
                    stats_text += f" | 总发行规模：<span style='font-weight:bold; color:#1e40af;'>{total_scale:,.2f} 万元</span>"

            # 预期收益率统计（保留原有）
            if "预期收益率" in df_projects.columns:
                valid_yields = df_projects[
                    df_projects["预期收益率"].apply(lambda x: pd.notna(x) and isinstance(x, (int, float)))]
                if not valid_yields.empty:
                    avg_yield = valid_yields["预期收益率"].mean()
                    stats_text += f" | 平均预期收益率：<span style='font-weight:bold; color:#1e40af;'>{avg_yield * 100:.2f}%</span>"

            st.markdown(f"""
            <div style="margin-top:10px; font-size:14px; color:#333;">
                {stats_text}
            </div>
            """, unsafe_allow_html=True)

        else:
            st.warning("⚠️ 当前筛选条件下无匹配项目，请调整关键词后重试")

    except FileNotFoundError:
        st.error("❌ 未找到项目库.xlsx文件，请确认文件在项目根目录（与app.py同级）")
    except Exception as e:
        st.error(f"❌ 项目清单加载失败：{str(e)}")
        st.info("请检查Excel文件格式（如Sheet名称、发行规模字段是否存在）")

    st.markdown("</div>", unsafe_allow_html=True)  # 项目清单卡片结束
    st.markdown("<br>", unsafe_allow_html=True)
