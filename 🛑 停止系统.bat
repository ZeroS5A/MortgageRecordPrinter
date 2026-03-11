@echo off
title 停止房贷信息录入系统
color 0C

echo ========================================
echo   正在关闭后台运行的房贷录入系统...
echo ========================================

:: 使用 WMIC 精准查找并结束运行了 streamlit 的 python 进程
:: 这样做可以避免误杀您电脑上正在运行的其他 Python 程序
wmic process where "name='python.exe' and commandline like '%%streamlit run app.py%%'" call terminate >nul 2>&1

echo.
echo ✅ 系统进程已安全结束！您可以关闭此窗口。
echo.

:: 停留 3 秒后自动关闭窗口
ping 127.0.0.1 -n 4 >nul
exit