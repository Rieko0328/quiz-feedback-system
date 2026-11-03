@echo off
chcp 65001 >nul
title 课测反馈系统 · 一键生成
cd /d "%~dp0"

REM 优先用自带的运行时，没有就用系统 python
set PY=
if exist "程序\python\python.exe" set PY=程序\python\python.exe
if "%PY%"=="" (
  where python >nul 2>&1 && set PY=python
)
if "%PY%"=="" (
  echo.
  echo   [错] 找不到 python。
  echo   要么把运行时放到 程序\python\ ，要么在系统里装一个 Python 3.10+。
  echo.
  pause
  exit /b 1
)

echo ================================================================
echo   课测反馈系统 · 一键生成
echo ================================================================
echo.
echo   用法：
echo     直接双击          = 让你输入班级
echo     拖班级名进来       = 直接跑那个班
echo.
echo   可用的班级：
for %%f in ("班级名单\*.txt") do (
  if not "%%~nf"=="_模板" echo       %%~nf
)
echo.

if "%~1"=="" (
  set /p CLS=  请输入班级名（照上面抄）：
) else (
  set CLS=%~1
)

if not exist "班级名单\%CLS%.txt" (
  echo.
  echo   [错] 找不到 班级名单\%CLS%.txt
  echo.
  pause
  exit /b 1
)

echo.
echo   正在处理：%CLS%
echo ----------------------------------------------------------------
"%PY%" "程序\脚本\一键生成.py" --班级 "%CLS%"
set RC=%ERRORLEVEL%
echo ----------------------------------------------------------------
if %RC%==0 (
  echo   ✔ 完成。产物在 生成结果\%CLS%_*  和  群发\%CLS%\
) else (
  echo   ✘ 出错了（退出码 %RC%），把上面红色的字发给技术看看。
)
echo.
pause
