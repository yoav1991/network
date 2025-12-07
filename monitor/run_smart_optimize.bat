@echo off
chcp 65001 >nul 2>&1
title Windows 网络性能智能优化工具

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 检查 Python
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo [错误] 未检测到 Python
    echo 请安装 Python 3.6+
    pause
    exit /b 1
)

python "%~dp0smart_optimize.py"
pause
