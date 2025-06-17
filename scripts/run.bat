@echo off
REM ================================================
REM Windows 批处理脚本：
REM 记录 run.bat 所在目录的上一级目录，并在该目录执行 uv run .\main.py
REM ================================================

chcp 65001

REM 获取 run.bat 所在目录
set "SCRIPT_DIR=%~dp0"

REM 去掉末尾的反斜杠（防止后面拼接）
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

REM 记录其上一级目录
for %%i in ("%SCRIPT_DIR%\..") do set "PARENT_DIR=%%~fi"

echo ================================================
echo run.bat 所在目录：%SCRIPT_DIR%
echo 上一级目录：%PARENT_DIR%
echo ================================================

REM 切换到上一级目录
cd /d "%PARENT_DIR%"

REM 执行 uv run .\main.py
echo 正在使用 uv 运行 main.py ...
uv run .\main.py

echo.
echo ✅ 执行完成，当前目录：%cd%
pause
@echo off
REM ================================================
REM Windows 批处理脚本：切换到上一级目录并使用 uv 运行 main.py
REM ================================================

chcp 65001

REM 记录当前目录
set "CURRENT_DIR=%cd%"
echo 当前目录：%CURRENT_DIR%

REM 切换到上一级目录
cd ..

echo 已切换到上一级目录：%cd%

REM 使用 uv 运行 main.py
echo 正在使用 uv 运行 main.py ...
uv run .\main.py

echo.f""*-0.000
echo ✅ 执行完成，当前仍在上一级目录：%cd%
pause
