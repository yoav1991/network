@echo off
chcp 65001 >nul 2>&1
title 网络快速修复工具

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo           Windows 网络快速修复工具
echo     专治: 浏览器无网络但 QQ/Telegram 正常的问题
echo ============================================================
echo.

echo [1/6] 正在禁用系统代理...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable /t REG_DWORD /d 0 /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyServer /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v AutoConfigURL /f >nul 2>&1
echo       完成

echo [2/6] 正在重置 WinHTTP 代理...
netsh winhttp reset proxy >nul 2>&1
echo       完成

echo [3/6] 正在刷新 DNS 缓存...
ipconfig /flushdns >nul 2>&1
echo       完成

echo [4/6] 正在重新注册 DNS...
ipconfig /registerdns >nul 2>&1
echo       完成

echo [5/6] 正在重置 Winsock...
netsh winsock reset >nul 2>&1
echo       完成

echo [6/6] 正在重置 TCP/IP 栈...
netsh int ip reset >nul 2>&1
echo       完成

echo.
echo ============================================================
echo                     修复操作已完成！
echo ============================================================
echo.
echo 注意: Winsock 和 TCP/IP 重置需要重启计算机才能完全生效
echo.

set /p restart="是否立即重启计算机? (Y/N): "
if /i "%restart%"=="Y" (
    echo 正在重启...
    shutdown /r /t 5
) else (
    echo.
    echo 请稍后手动重启计算机以完成修复
)

pause
