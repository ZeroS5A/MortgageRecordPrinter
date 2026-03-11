Set ws = CreateObject("WScript.Shell")
' 运行 streamlit 命令，0 表示隐藏黑框，False 表示不阻塞直接返回
ws.Run "cmd /c chcp 65001 && set PYTHONIOENCODING=utf-8 && streamlit run app.py", 0, False