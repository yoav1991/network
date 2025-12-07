@echo off
chcp 65001 >nul 2>&1
title Windows 网络监控与优化工具 (PowerShell版)

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: 运行 PowerShell 脚本
powershell -ExecutionPolicy Bypass -File "%~dp0NetworkMonitor.ps1"

pause
