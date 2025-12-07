@echo off
chcp 65001 >nul 2>&1
title 网络快速优化

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo           网络快速优化工具 - 解决网页变慢问题
echo ============================================================
echo.

echo [1/5] 正在清理 DNS 缓存...
ipconfig /flushdns >nul 2>&1
echo       完成

echo [2/5] 正在清理 ARP 缓存...
netsh interface ip delete arpcache >nul 2>&1
arp -d * >nul 2>&1
echo       完成

echo [3/5] 正在清理 NetBIOS 缓存...
nbtstat -R >nul 2>&1
echo       完成

echo [4/5] 正在刷新 NetBIOS 会话...
nbtstat -RR >nul 2>&1
echo       完成

echo [5/5] 正在重新注册 DNS...
ipconfig /registerdns >nul 2>&1
echo       完成

echo.
echo ============================================================
echo                  优化完成！网络缓存已清理
echo ============================================================
echo.
echo 建议：如果问题持续，可以设置定时优化任务
echo       运行 setup_scheduled_task.bat 创建定时任务
echo.

pause
