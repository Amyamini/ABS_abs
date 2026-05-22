@echo off
chcp 65001 >nul
echo ========================================
echo   修复 page_product_liquidity.py 缩进
echo ========================================
echo.

python quick_fix.py

echo.
echo 按任意键退出...
pause >nul
