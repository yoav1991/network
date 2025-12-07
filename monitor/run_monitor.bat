@echo off
chcp 65001 >nul 2>&1
title Windows 网络监控与优化工具

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 检查 Python 是否安装
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo [错误] 未检测到 Python 安装
    echo 请先安装 Python 3.6 或更高版本
    echo 下载地址: https://www.python.org/downloads/
    echo.
    echo 或者使用 PowerShell 版本: run_monitor_ps.bat
    echo.
    pause
    exit /b 1
)

:: 运行监控工具
python "%~dp0network_monitor.py"

pause
