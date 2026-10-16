@echo off
chcp 65001 >nul
title 课测反馈系统
cd /d "%~dp0"
"程序\python\python.exe" "程序\app.py"
