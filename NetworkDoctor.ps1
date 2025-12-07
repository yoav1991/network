#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows 网络诊断修复工具 (PowerShell 版本)

.DESCRIPTION
    用于检测和修复 Windows 系统中浏览器无网络但通讯软件正常的问题

.NOTES
    版本: 1.0.0
    需要管理员权限运行
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Windows 网络诊断修复工具"

# 颜色定义
function Write-ColorText {
    param(
        [string]$Text,
        [string]$Color = "White"
    )
    Write-Host $Text -ForegroundColor $Color -NoNewline
}

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "      Windows 网络诊断修复工具 (Network Doctor) v1.0" -ForegroundColor White
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  专门解决: 浏览器无网络但QQ/Telegram等通讯软件正常的问题" -ForegroundColor Yellow
    Write-Host ""
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "──────────────────────────────────────────────────" -ForegroundColor Blue
    Write-Host "【$Title】" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────────" -ForegroundColor Blue
}

# 诊断结果类
class DiagResult {
    [string]$Name
    [string]$Status  # OK, WARNING, ERROR, INFO
    [string]$Message
    [string]$Details
    [bool]$FixAvailable
}

# ============== 诊断函数 ==============

function Test-ProxySettings {
    Write-Host "  → 检查系统代理设置..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "系统代理设置 (IE/Edge)"

    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        $proxyEnable = (Get-ItemProperty -Path $regPath -Name ProxyEnable -ErrorAction SilentlyContinue).ProxyEnable
        $proxyServer = (Get-ItemProperty -Path $regPath -Name ProxyServer -ErrorAction SilentlyContinue).ProxyServer
        $autoConfigURL = (Get-ItemProperty -Path $regPath -Name AutoConfigURL -ErrorAction SilentlyContinue).AutoConfigURL

        $result.Details = "代理启用: $(if($proxyEnable){'是'}else{'否'})`n"
        $result.Details += "代理服务器: $(if($proxyServer){$proxyServer}else{'无'})`n"
        $result.Details += "PAC脚本: $(if($autoConfigURL){$autoConfigURL}else{'无'})"

        if ($proxyEnable -and $proxyServer) {
            $result.Status = "WARNING"
            $result.Message = "检测到系统代理已启用，这可能导致部分程序无法联网"
            $result.FixAvailable = $true
        }
        elseif ($autoConfigURL) {
            $result.Status = "INFO"
            $result.Message = "检测到 PAC 自动配置脚本"
            $result.FixAvailable = $true
        }
        else {
            $result.Status = "OK"
            $result.Message = "系统代理未启用"
        }
    }
    catch {
        $result.Status = "ERROR"
        $result.Message = "无法读取代理设置: $_"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-WinHTTPProxy {
    Write-Host "  → 检查 WinHTTP 代理..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "WinHTTP 代理设置"

    try {
        $output = netsh winhttp show proxy 2>&1
        $result.Details = $output -join "`n"

        if ($output -match "直接访问|Direct access") {
            $result.Status = "OK"
            $result.Message = "WinHTTP 设置为直接访问（无代理）"
        }
        elseif ($output -match "代理服务器|Proxy Server") {
            $result.Status = "WARNING"
            $result.Message = "检测到 WinHTTP 代理设置"
            $result.FixAvailable = $true
        }
        else {
            $result.Status = "INFO"
            $result.Message = "WinHTTP 代理状态"
        }
    }
    catch {
        $result.Status = "ERROR"
        $result.Message = "无法获取 WinHTTP 设置: $_"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-DNSSettings {
    Write-Host "  → 检查 DNS 设置..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "DNS 设置"

    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        $dnsServers = @()

        foreach ($adapter in $adapters) {
            $dns = Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
            if ($dns.ServerAddresses) {
                $dnsServers += $dns.ServerAddresses
            }
        }

        $dnsServers = $dnsServers | Select-Object -Unique
        $result.Details = "检测到的 DNS 服务器: $($dnsServers -join ', ')"

        if ($dnsServers.Count -eq 0) {
            $result.Status = "ERROR"
            $result.Message = "未检测到 DNS 服务器配置"
            $result.FixAvailable = $true
        }
        else {
            $result.Status = "OK"
            $result.Message = "已配置 $($dnsServers.Count) 个 DNS 服务器"
        }
    }
    catch {
        $result.Status = "ERROR"
        $result.Message = "无法获取 DNS 设置: $_"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-DNSResolution {
    Write-Host "  → 测试 DNS 解析..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "DNS 解析测试"

    $testDomains = @(
        @{Domain="www.baidu.com"; Name="百度"},
        @{Domain="www.qq.com"; Name="腾讯"},
        @{Domain="www.microsoft.com"; Name="微软"}
    )

    $results = @()
    $failed = 0

    foreach ($test in $testDomains) {
        try {
            $ip = [System.Net.Dns]::GetHostAddresses($test.Domain)[0].IPAddressToString
            $results += "$($test.Name)($($test.Domain)): $ip"
        }
        catch {
            $results += "$($test.Name)($($test.Domain)): 解析失败"
            $failed++
        }
    }

    $result.Details = $results -join "`n"

    if ($failed -eq $testDomains.Count) {
        $result.Status = "ERROR"
        $result.Message = "DNS 解析完全失败，这可能是主要问题"
        $result.FixAvailable = $true
    }
    elseif ($failed -gt 0) {
        $result.Status = "WARNING"
        $result.Message = "部分域名解析失败 ($failed/$($testDomains.Count))"
        $result.FixAvailable = $true
    }
    else {
        $result.Status = "OK"
        $result.Message = "DNS 解析正常"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-NetworkConnectivity {
    Write-Host "  → 测试网络连通性..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "网络连通性"

    $testTargets = @(
        @{IP="114.114.114.114"; Name="国内DNS"},
        @{IP="223.5.5.5"; Name="阿里DNS"},
        @{IP="8.8.8.8"; Name="GoogleDNS"}
    )

    $results = @()
    $success = 0

    foreach ($target in $testTargets) {
        $ping = Test-Connection -ComputerName $target.IP -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($ping) {
            $results += "$($target.Name)($($target.IP)): 可达"
            $success++
        }
        else {
            $results += "$($target.Name)($($target.IP)): 不可达"
        }
    }

    $result.Details = $results -join "`n"

    if ($success -eq 0) {
        $result.Status = "ERROR"
        $result.Message = "无法连接到任何测试目标"
    }
    elseif ($success -lt $testTargets.Count) {
        $result.Status = "WARNING"
        $result.Message = "部分目标可达 ($success/$($testTargets.Count))"
    }
    else {
        $result.Status = "OK"
        $result.Message = "网络连通性正常"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-HTTPConnectivity {
    Write-Host "  → 测试 HTTP 连通性..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "HTTP 连通性测试"

    $testURLs = @(
        @{URL="http://www.baidu.com"; Name="百度HTTP"},
        @{URL="https://www.qq.com"; Name="腾讯HTTPS"}
    )

    $results = @()
    $success = 0

    # 禁用 SSL 证书验证（仅用于测试）
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

    foreach ($test in $testURLs) {
        try {
            $response = Invoke-WebRequest -Uri $test.URL -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
            if ($response.StatusCode -eq 200) {
                $results += "$($test.Name): 正常 (状态码: $($response.StatusCode))"
                $success++
            }
            else {
                $results += "$($test.Name): 异常 (状态码: $($response.StatusCode))"
            }
        }
        catch {
            $results += "$($test.Name): 失败 ($($_.Exception.Message.Substring(0, [Math]::Min(30, $_.Exception.Message.Length))))"
        }
    }

    $result.Details = $results -join "`n"

    if ($success -eq 0) {
        $result.Status = "ERROR"
        $result.Message = "HTTP 请求全部失败！这是问题的关键表现"
        $result.FixAvailable = $true
    }
    elseif ($success -lt $testURLs.Count) {
        $result.Status = "WARNING"
        $result.Message = "部分 HTTP 请求失败"
        $result.FixAvailable = $true
    }
    else {
        $result.Status = "OK"
        $result.Message = "HTTP 连通性正常"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-HostsFile {
    Write-Host "  → 检查 Hosts 文件..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "Hosts 文件"

    try {
        $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
        $content = Get-Content $hostsPath -ErrorAction Stop
        $lines = $content | Where-Object { $_ -and -not $_.StartsWith('#') }

        $suspicious = @()
        $keywords = @('google', 'facebook', 'youtube', 'twitter', 'github')

        foreach ($line in $lines) {
            foreach ($keyword in $keywords) {
                if ($line -match $keyword) {
                    $suspicious += $line
                    break
                }
            }
        }

        $result.Details = "自定义条目数: $($lines.Count)"
        if ($suspicious) {
            $result.Details += "`n可疑条目:`n$($suspicious[0..4] -join "`n")"
        }

        if ($suspicious.Count -gt 0) {
            $result.Status = "WARNING"
            $result.Message = "检测到 $($suspicious.Count) 个可能影响网络的 hosts 条目"
            $result.FixAvailable = $true
        }
        else {
            $result.Status = "OK"
            $result.Message = "Hosts 文件正常"
        }
    }
    catch {
        $result.Status = "ERROR"
        $result.Message = "无法读取 hosts 文件: $_"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

function Test-Winsock {
    Write-Host "  → 检查 Winsock 目录..." -ForegroundColor Cyan -NoNewline

    $result = [DiagResult]::new()
    $result.Name = "Winsock 目录"

    try {
        $output = netsh winsock show catalog 2>&1
        $outputText = $output -join " "

        $suspiciousKeywords = @('proxy', 'vpn', 'hook', 'inject', 'filter')
        $found = @()

        foreach ($keyword in $suspiciousKeywords) {
            if ($outputText -match $keyword) {
                $found += $keyword
            }
        }

        if ($found.Count -gt 0) {
            $result.Status = "WARNING"
            $result.Message = "检测到可能影响网络的第三方 LSP: $($found -join ', ')"
            $result.FixAvailable = $true
        }
        else {
            $result.Status = "OK"
            $result.Message = "Winsock 目录看起来正常"
        }

        $result.Details = "检查完成"
    }
    catch {
        $result.Status = "ERROR"
        $result.Message = "无法获取 Winsock 信息: $_"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $result
}

# ============== 修复函数 ==============

function Repair-DisableProxy {
    Write-Host "  → 正在禁用系统代理..." -ForegroundColor Cyan

    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 0 -ErrorAction Stop
        Remove-ItemProperty -Path $regPath -Name ProxyServer -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $regPath -Name AutoConfigURL -ErrorAction SilentlyContinue

        Write-Host "    ✓ 系统代理已禁用" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 禁用失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-ResetWinHTTP {
    Write-Host "  → 正在重置 WinHTTP 代理..." -ForegroundColor Cyan

    try {
        $null = netsh winhttp reset proxy 2>&1
        Write-Host "    ✓ WinHTTP 代理已重置" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 重置失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-FlushDNS {
    Write-Host "  → 正在刷新 DNS 缓存..." -ForegroundColor Cyan

    try {
        $null = ipconfig /flushdns 2>&1
        Write-Host "    ✓ DNS 缓存已刷新" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 刷新失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-RegisterDNS {
    Write-Host "  → 正在重新注册 DNS..." -ForegroundColor Cyan

    try {
        $null = ipconfig /registerdns 2>&1
        Write-Host "    ✓ DNS 已重新注册" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 注册失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-ResetWinsock {
    Write-Host "  → 正在重置 Winsock..." -ForegroundColor Cyan

    try {
        $null = netsh winsock reset 2>&1
        Write-Host "    ✓ Winsock 已重置 (需要重启生效)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 重置失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-ResetTCPIP {
    Write-Host "  → 正在重置 TCP/IP 栈..." -ForegroundColor Cyan

    try {
        $null = netsh int ip reset 2>&1
        Write-Host "    ✓ TCP/IP 栈已重置 (需要重启生效)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 重置失败: $_" -ForegroundColor Red
        return $false
    }
}

function Repair-ReleaseRenewIP {
    Write-Host "  → 正在释放/更新 IP 地址..." -ForegroundColor Cyan

    try {
        $null = ipconfig /release 2>&1
        Start-Sleep -Seconds 2
        $null = ipconfig /renew 2>&1
        Write-Host "    ✓ IP 地址已更新" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 更新失败: $_" -ForegroundColor Red
        return $false
    }
}

# ============== 主程序 ==============

function Show-DiagnosticResult {
    param([DiagResult]$Result)

    $statusColor = switch ($Result.Status) {
        "OK" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "INFO" { "Blue" }
        default { "White" }
    }

    $statusIcon = switch ($Result.Status) {
        "OK" { "✓ 正常" }
        "WARNING" { "⚠ 警告" }
        "ERROR" { "✗ 异常" }
        "INFO" { "ℹ 信息" }
        default { "?" }
    }

    Write-Host ""
    Write-Host "  " -NoNewline
    Write-Host $statusIcon -ForegroundColor $statusColor -NoNewline
    Write-Host " $($Result.Name)"
    Write-Host "      $($Result.Message)"

    if ($Result.Details) {
        foreach ($line in $Result.Details.Split("`n")) {
            if ($line.Trim()) {
                Write-Host "      $line" -ForegroundColor Gray
            }
        }
    }
}

function Run-AllDiagnostics {
    Write-Section "开始网络诊断"

    $results = @()
    $issuesFound = 0

    $results += Test-ProxySettings
    $results += Test-WinHTTPProxy
    $results += Test-Winsock
    $results += Test-DNSSettings
    $results += Test-DNSResolution
    $results += Test-HostsFile
    $results += Test-NetworkConnectivity
    $results += Test-HTTPConnectivity

    foreach ($r in $results) {
        if ($r.Status -eq "WARNING" -or $r.Status -eq "ERROR") {
            $issuesFound++
        }
    }

    Write-Section "诊断结果摘要"

    foreach ($r in $results) {
        Show-DiagnosticResult -Result $r
    }

    Write-Host ""
    Write-Host "──────────────────────────────────────────────────" -ForegroundColor Blue

    if ($issuesFound -gt 0) {
        Write-Host "  发现 $issuesFound 个潜在问题" -ForegroundColor Yellow
        Write-Host "  建议执行修复操作来解决网络问题"
    }
    else {
        Write-Host "  未发现明显问题" -ForegroundColor Green
        Write-Host "  如果问题持续，建议尝试完整修复"
    }

    return $results
}

function Run-QuickRepair {
    Write-Section "开始快速修复"

    Repair-DisableProxy
    Repair-ResetWinHTTP
    Repair-FlushDNS

    Write-Host ""
    Write-Host "快速修复完成！" -ForegroundColor Green
    Write-Host "如果问题仍然存在，请尝试完整修复" -ForegroundColor Yellow
}

function Run-FullRepair {
    Write-Section "开始完整修复"

    Repair-DisableProxy
    Repair-ResetWinHTTP
    Repair-FlushDNS
    Repair-RegisterDNS
    Repair-ResetWinsock
    Repair-ResetTCPIP
    Repair-ReleaseRenewIP

    Write-Host ""
    Write-Host "完整修复完成！" -ForegroundColor Green
    Write-Host "某些更改需要重启计算机才能生效" -ForegroundColor Yellow

    $restart = Read-Host "`n是否立即重启计算机? (Y/N)"
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-Host "正在重启..."
        Restart-Computer -Force
    }
}

function Show-Menu {
    Write-Host ""
    Write-Host "──────────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "请选择操作:" -ForegroundColor White
    Write-Host "  1. 运行网络诊断"
    Write-Host "  2. 快速修复 (推荐首先尝试)"
    Write-Host "  3. 完整修复 (需要重启)"
    Write-Host "  4. 诊断 + 自动修复"
    Write-Host "  5. 仅禁用系统代理"
    Write-Host "  6. 仅重置 Winsock"
    Write-Host "  0. 退出"
    Write-Host "──────────────────────────────────────────────────" -ForegroundColor Cyan

    return Read-Host "`n请输入选项 [0-6]"
}

# 主循环
Write-Header

while ($true) {
    $choice = Show-Menu

    switch ($choice) {
        '0' {
            Write-Host "`n感谢使用网络诊断修复工具！`n" -ForegroundColor Cyan
            exit
        }
        '1' {
            $null = Run-AllDiagnostics
        }
        '2' {
            Run-QuickRepair
        }
        '3' {
            Run-FullRepair
        }
        '4' {
            $results = Run-AllDiagnostics
            $issues = ($results | Where-Object { $_.Status -eq "WARNING" -or $_.Status -eq "ERROR" }).Count

            if ($issues -gt 0) {
                Write-Host "`n检测到问题，是否进行自动修复?" -ForegroundColor Yellow
                $confirm = Read-Host "输入 Y 确认修复，其他键跳过"
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Run-FullRepair
                }
            }
        }
        '5' {
            Write-Section "禁用代理"
            Repair-DisableProxy
            Repair-ResetWinHTTP
            Write-Host "`n代理设置已清除！" -ForegroundColor Green
        }
        '6' {
            Write-Section "重置 Winsock"
            Repair-ResetWinsock
            Write-Host "`nWinsock 已重置，需要重启生效" -ForegroundColor Green
        }
        default {
            Write-Host "`n无效选项，请重新选择" -ForegroundColor Red
        }
    }
}
