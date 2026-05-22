@echo off
chcp 65001 >nul
echo ========================================
echo   修复页面缩进
echo ========================================
echo.

python fix_indent_for_menu.py

echo.
echo 按任意键退出...
pause >nul
