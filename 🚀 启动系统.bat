@echo off
:: ==========================================
:: 隐藏 CMD 窗口逻辑 (核心修改)
:: ==========================================
if "%1" == "h" goto begin
:: 利用 mshta 调用 vbscript 隐藏运行自身
mshta vbscript:createobject("wscript.shell").run("""%~f0"" h",0)(window.close)&&exit
:begin

setlocal

:: ==========================================
:: 基础配置
:: ==========================================

@echo off
chcp 65001 >nul
set PYTHONUTF8=1
streamlit run app.py --server.port 8502