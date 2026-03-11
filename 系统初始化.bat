@echo off
title 房贷信息录入系统启动程序
color 0A

echo ===================================================
echo   正在检查并安装必要的运行库，请保持网络畅通...
echo ===================================================
pip install streamlit openpyxl pandas pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo 环境检查完毕，正在启动系统...
echo (不要关闭此黑色窗口，请在自动弹出的浏览器中进行操作)
echo.

streamlit run app.py

pause