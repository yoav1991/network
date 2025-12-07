@echo off
chcp 65001 >nul 2>&1
title Windows 网络性能智能优化工具

:: 检查管理员权限
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 正在请求管理员权限...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:: 检查 Python
where python >nul 2>&1
if '%errorlevel%' NEQ '0' (
    echo.
    echo [错误] 未检测到 Python
    echo 请安装 Python 3.6+
    echo.
    pause
    exit /B 1
)

:: 运行脚本
echo 正在启动智能优化工具...
python "%~dp0smart_optimize.py"

echo.
pause
