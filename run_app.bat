@echo off
chcp 65001 >nul
echo ========================================
echo   ABS投资跟进系统 - 启动脚本
echo ========================================
echo.
echo 正在启动应用...
echo.
echo 请在浏览器中访问: http://localhost:8501
echo.
echo 按 Ctrl+C 可停止应用
echo.

streamlit run app.py

pause
