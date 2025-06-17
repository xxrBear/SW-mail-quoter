@echo off
REM ================================================
REM Windows 批处理脚本：安装 uv（若未安装）并执行 uv sync
REM ================================================

REM 设置终端为 UTF-8 编码
chcp 65001

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
@REM pause

echo 正在写入 .env 到上一层目录...

REM %~dp0 是当前批处理所在的绝对路径
REM 用 for 循环解析真正的绝对上层目录
for %%i in ("%~dp0..") do set PARENT_DIR=%%~fi

REM 检查
echo 上层目录: %PARENT_DIR%

REM 写入文件
echo EMAIL_SMTP_SERVER='' > "%PARENT_DIR%\.env"
echo EMAIL_USER_NAME='' >> "%PARENT_DIR%\.env"
echo EMAIL_USER_PASS='' >> "%PARENT_DIR%\.env"

echo.
echo 已在上一层目录生成 .env 文件，内容如下：
type "%PARENT_DIR%\.env"echo 正在写入 .env 到上一层目录...

pause
