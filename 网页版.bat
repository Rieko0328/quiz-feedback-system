@echo off
chcp 65001 >nul
title 课测反馈系统 · 网页版
cd /d "%~dp0"

echo.
echo   ============================================
echo      课测反馈系统  ·  网页版
echo   ============================================
echo.
echo   正在启动，浏览器会自动打开...
echo   这个黑窗口不要关，关了系统就停了。
echo.

"程序\python\python.exe" "程序\web_server.py" --端口 8777
if errorlevel 1 (
    echo.
    echo   [!] 启动失败或端口被占用，换一个端口再试...
    echo.
    "程序\python\python.exe" "程序\web_server.py" --端口 8788
)
echo.
echo   系统已停止。按任意键关闭窗口。
pause >nul
