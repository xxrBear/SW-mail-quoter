@echo off
REM ================================================
REM Windows 批处理脚本：切换到上一级目录并执行 uv run .\main.py
REM ================================================

chcp 65001

echo 正在切换到上一级目录...
cd ..

echo 当前目录：%cd%
echo 正在使用 uv 运行 main.py...

uv run .\main.py

echo.
pause
