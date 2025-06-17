@echo off
REM ================================================
REM Windows 批处理脚本：安装 uv（若未安装）并执行 uv sync
REM 然后仅在上一层目录没有 .env 时才创建 .env
REM ================================================

chcp 65001

REM -----------------------------------------------
REM 检查 uv 是否已安装
where uv >nul 2>nul

if %ERRORLEVEL% EQU 0 (
    echo 已检测到 uv，版本信息如下：
    uv --version
) else (
    echo 未检测到 uv，正在安装 uv ...
    powershell -ExecutionPolicy ByPass -Command "irm https://astral.sh/uv/install.ps1 | iex"

    REM 再次检查是否安装成功
    where uv >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        echo uv 安装成功！
    ) else (
        echo uv 安装失败，请检查网络或重试！
        pause
        exit /b 1
    )
)

echo.
echo ================================================
echo 正在执行 uv sync ...
uv sync

echo.
echo uv sync 执行完成！
echo ================================================

REM -----------------------------------------------
REM 记录批处理所在目录的上一级目录
for %%i in ("%~dp0..") do set "PARENT_DIR=%%~fi"

REM 检查 .env 是否存在
if exist "%PARENT_DIR%\.env" (
    echo 已检测到上一层目录已存在 .env 文件，跳过创建。
) else (
    echo 上一层目录未发现 .env 文件，正在创建...

    echo EMAIL_SMTP_SERVER='' > "%PARENT_DIR%\.env"
    echo EMAIL_USER_NAME='' >> "%PARENT_DIR%\.env"
    echo EMAIL_USER_PASS='' >> "%PARENT_DIR%\.env"

    echo 已成功在上一层目录生成 .env 文件，内容如下：
    type "%PARENT_DIR%\.env"
)

pause
