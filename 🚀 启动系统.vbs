Set ws = CreateObject("WScript.Shell")
' 0 表示隐藏命令行黑框，False 表示不阻塞直接返回
' 这里也加入了编码设置，确保后台静默运行时也能正确处理中文字符
ws.Run "cmd /c chcp 65001 & set PYTHONUTF8=1 & streamlit run app.py", 0, False