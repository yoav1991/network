@echo off
chcp 65001 >nul 2>&1
title 设置定时网络优化任务

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo              设置定时网络优化任务
echo ============================================================
echo.
echo 此脚本将创建一个 Windows 计划任务，
echo 每隔一定时间自动清理网络缓存，防止系统变慢。
echo.
echo 请选择优化频率:
echo   1. 每 30 分钟 (推荐长时间运行的电脑)
echo   2. 每 1 小时
echo   3. 每 2 小时
echo   4. 每 4 小时
echo   5. 删除已有的定时任务
echo   0. 取消
echo.

set /p choice="请输入选项 [0-5]: "

if "%choice%"=="0" goto :cancel
if "%choice%"=="5" goto :delete

:: 创建优化脚本
set SCRIPT_PATH=%~dp0auto_optimize.bat
echo @echo off > "%SCRIPT_PATH%"
echo chcp 65001 ^>nul 2^>^&1 >> "%SCRIPT_PATH%"
echo ipconfig /flushdns ^>nul 2^>^&1 >> "%SCRIPT_PATH%"
echo netsh interface ip delete arpcache ^>nul 2^>^&1 >> "%SCRIPT_PATH%"
echo arp -d * ^>nul 2^>^&1 >> "%SCRIPT_PATH%"
echo nbtstat -R ^>nul 2^>^&1 >> "%SCRIPT_PATH%"
echo nbtstat -RR ^>nul 2^>^&1 >> "%SCRIPT_PATH%"

:: 删除已有任务
schtasks /delete /tn "NetworkOptimizer" /f >nul 2>&1

:: 根据选择设置间隔
if "%choice%"=="1" set INTERVAL=30
if "%choice%"=="2" set INTERVAL=60
if "%choice%"=="3" set INTERVAL=120
if "%choice%"=="4" set INTERVAL=240

:: 创建计划任务
schtasks /create /tn "NetworkOptimizer" /tr "\"%SCRIPT_PATH%\"" /sc minute /mo %INTERVAL% /ru SYSTEM /rl HIGHEST /f

if %errorLevel% equ 0 (
    echo.
    echo ============================================================
    echo                 定时任务创建成功！
    echo ============================================================
    echo.
    echo 任务名称: NetworkOptimizer
    echo 执行间隔: 每 %INTERVAL% 分钟
    echo 任务路径: %SCRIPT_PATH%
    echo.
    echo 系统将自动在后台清理网络缓存，无需手动操作。
    echo.
) else (
    echo.
    echo 创建任务失败，请确保以管理员身份运行。
)
goto :end

:delete
schtasks /delete /tn "NetworkOptimizer" /f >nul 2>&1
if exist "%~dp0auto_optimize.bat" del "%~dp0auto_optimize.bat"
echo.
echo 定时任务已删除。
goto :end

:cancel
echo.
echo 已取消。

:end
echo.
pause
