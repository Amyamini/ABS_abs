@echo off
chcp 65001 >nul
echo ========================================
echo   启动产品流动性管理系统
echo   （包含母子基金联动功能）
echo ========================================
echo.
echo 正在启动 Streamlit 应用...
echo.
streamlit run page_product_liquidity.py
pause
