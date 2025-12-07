@echo off
chcp 65001 >nul 2>&1
title 一键禁用代理

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo                   一键禁用所有代理设置
echo ============================================================
echo.

echo [1/2] 正在禁用 IE/系统代理...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable /t REG_DWORD /d 0 /f
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyServer /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v AutoConfigURL /f >nul 2>&1
echo       完成

echo [2/2] 正在重置 WinHTTP 代理...
netsh winhttp reset proxy
echo       完成

echo.
echo ============================================================
echo               代理已禁用！请测试网络连接
echo ============================================================
echo.

pause
