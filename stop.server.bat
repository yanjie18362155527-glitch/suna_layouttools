@echo off
echo 正在寻找占用 8501 端口的 Streamlit 服务...

:: 查找占用 8501 端口的进程 PID，并强制终止它
for /f "tokens=5" %%a in ('netstat -aon ^| find ":8501" ^| find "LISTENING"') do (
    echo 发现进程 PID: %%a，正在终止...
    taskkill /f /pid %%a
)

echo.
echo 服务已停止！
pause