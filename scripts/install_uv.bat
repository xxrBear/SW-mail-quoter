@echo off
REM ================================================
REM Windows 批处理脚本：安装 uv（若未安装）
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
        powershell -Command "Write-Host '按任意键继续 . . .' -NoNewline; $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
        exit /b 1
    )
)
