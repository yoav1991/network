#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows 网络监控与优化工具 (PowerShell 版本)

.DESCRIPTION
    用于实时监控网络状态并定期优化，解决长时间运行后网速变慢的问题

.NOTES
    版本: 1.0.0
    需要管理员权限运行
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Windows 网络监控与优化工具"

# ============== 监控函数 ==============

function Measure-DNSResolveTime {
    param([string]$Domain = "www.baidu.com")

    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $null = [System.Net.Dns]::GetHostAddresses($Domain)
        $stopwatch.Stop()
        return $stopwatch.ElapsedMilliseconds
    }
    catch {
        return -1
    }
}

function Measure-PingTime {
    param([string]$Host = "114.114.114.114")

    try {
        $ping = Test-Connection -ComputerName $Host -Count 1 -ErrorAction Stop
        return $ping.ResponseTime
    }
    catch {
        return -1
    }
}

function Measure-HTTPResponseTime {
    param([string]$Url = "http://www.baidu.com")

    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $response = Invoke-WebRequest -Uri $Url -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
        $stopwatch.Stop()
        return $stopwatch.ElapsedMilliseconds
    }
    catch {
        return -1
    }
}

function Get-TCPConnectionCount {
    try {
        $connections = Get-NetTCPConnection -ErrorAction SilentlyContinue
        return $connections.Count
    }
    catch {
        return -1
    }
}

function Get-DNSCacheCount {
    try {
        $cache = Get-DnsClientCache -ErrorAction SilentlyContinue
        return $cache.Count
    }
    catch {
        return -1
    }
}

function Get-ARPCacheCount {
    try {
        $arp = Get-NetNeighbor -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Permanent' }
        return $arp.Count
    }
    catch {
        return -1
    }
}

function Get-MemoryUsagePercent {
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $usedMemory = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
        return [math]::Round(($usedMemory / $os.TotalVisibleMemorySize) * 100, 1)
    }
    catch {
        return -1
    }
}

function Get-NetworkStats {
    Write-Host "  正在收集网络状态..." -ForegroundColor Cyan -NoNewline

    $stats = [PSCustomObject]@{
        Timestamp = Get-Date
        DNSResolveTime = Measure-DNSResolveTime
        PingTime = Measure-PingTime
        HTTPResponseTime = Measure-HTTPResponseTime
        TCPConnections = Get-TCPConnectionCount
        DNSCacheEntries = Get-DNSCacheCount
        ARPCacheEntries = Get-ARPCacheCount
        MemoryUsage = Get-MemoryUsagePercent
        Status = "正常"
    }

    # 判断状态
    if ($stats.HTTPResponseTime -gt 3000) {
        $stats.Status = "缓慢"
    }
    elseif ($stats.HTTPResponseTime -gt 1500) {
        $stats.Status = "较慢"
    }
    elseif ($stats.HTTPResponseTime -lt 0) {
        $stats.Status = "异常"
    }

    Write-Host " 完成" -ForegroundColor Green
    return $stats
}

function Show-NetworkStats {
    param($Stats)

    $statusColors = @{
        "正常" = "Green"
        "较慢" = "Yellow"
        "缓慢" = "Red"
        "异常" = "Red"
    }

    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ("  时间: {0}    状态: " -f $Stats.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")) -NoNewline
    Write-Host $Stats.Status -ForegroundColor $statusColors[$Stats.Status]
    Write-Host ("─" * 65) -ForegroundColor Blue

    # 网络延迟
    Write-Host "`n  网络延迟指标:" -ForegroundColor White

    $dnsColor = if ($Stats.DNSResolveTime -lt 100) { "Green" } elseif ($Stats.DNSResolveTime -lt 500) { "Yellow" } else { "Red" }
    $pingColor = if ($Stats.PingTime -lt 50) { "Green" } elseif ($Stats.PingTime -lt 100) { "Yellow" } else { "Red" }
    $httpColor = if ($Stats.HTTPResponseTime -lt 500) { "Green" } elseif ($Stats.HTTPResponseTime -lt 1500) { "Yellow" } else { "Red" }

    Write-Host ("    DNS 解析时间:   {0,8:N0} ms" -f $Stats.DNSResolveTime) -ForegroundColor $dnsColor
    Write-Host ("    Ping 延迟:      {0,8:N0} ms" -f $Stats.PingTime) -ForegroundColor $pingColor
    Write-Host ("    HTTP 响应时间:  {0,8:N0} ms" -f $Stats.HTTPResponseTime) -ForegroundColor $httpColor

    # 系统资源
    Write-Host "`n  系统资源状态:" -ForegroundColor White

    $tcpColor = if ($Stats.TCPConnections -lt 300) { "Green" } elseif ($Stats.TCPConnections -lt 500) { "Yellow" } else { "Red" }
    $dnsChColor = if ($Stats.DNSCacheEntries -lt 500) { "Green" } elseif ($Stats.DNSCacheEntries -lt 1000) { "Yellow" } else { "Red" }
    $arpColor = if ($Stats.ARPCacheEntries -lt 100) { "Green" } elseif ($Stats.ARPCacheEntries -lt 200) { "Yellow" } else { "Red" }
    $memColor = if ($Stats.MemoryUsage -lt 70) { "Green" } elseif ($Stats.MemoryUsage -lt 85) { "Yellow" } else { "Red" }

    Write-Host ("    TCP 连接数:     {0,8} 个" -f $Stats.TCPConnections) -ForegroundColor $tcpColor
    Write-Host ("    DNS 缓存条目:   {0,8} 条" -f $Stats.DNSCacheEntries) -ForegroundColor $dnsChColor
    Write-Host ("    ARP 缓存条目:   {0,8} 条" -f $Stats.ARPCacheEntries) -ForegroundColor $arpColor
    Write-Host ("    内存使用率:     {0,7:N1} %" -f $Stats.MemoryUsage) -ForegroundColor $memColor

    # 优化建议
    $suggestions = @()
    if ($Stats.DNSCacheEntries -gt 1000) { $suggestions += "DNS 缓存过大，建议清理" }
    if ($Stats.TCPConnections -gt 500) { $suggestions += "TCP 连接数过多，可能存在连接泄漏" }
    if ($Stats.ARPCacheEntries -gt 200) { $suggestions += "ARP 缓存过大，建议清理" }
    if ($Stats.HTTPResponseTime -gt 2000) { $suggestions += "网络响应缓慢，建议执行优化" }

    if ($suggestions.Count -gt 0) {
        Write-Host "`n  优化建议:" -ForegroundColor Yellow
        foreach ($s in $suggestions) {
            Write-Host ("    → " + $s) -ForegroundColor Yellow
        }
    }
}

# ============== 优化函数 ==============

function Clear-DNSCache {
    Write-Host "  → 正在清理 DNS 缓存..." -ForegroundColor Cyan
    try {
        Clear-DnsClientCache -ErrorAction Stop
        Write-Host "    ✓ DNS 缓存已清理" -ForegroundColor Green
        return $true
    }
    catch {
        $null = ipconfig /flushdns 2>&1
        Write-Host "    ✓ DNS 缓存已清理 (备用方法)" -ForegroundColor Green
        return $true
    }
}

function Clear-ARPCache {
    Write-Host "  → 正在清理 ARP 缓存..." -ForegroundColor Cyan
    try {
        $null = netsh interface ip delete arpcache 2>&1
        $null = arp -d * 2>&1
        Write-Host "    ✓ ARP 缓存已清理" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 清理失败" -ForegroundColor Red
        return $false
    }
}

function Clear-NetBIOSCache {
    Write-Host "  → 正在清理 NetBIOS 缓存..." -ForegroundColor Cyan
    try {
        $null = nbtstat -R 2>&1
        Write-Host "    ✓ NetBIOS 缓存已清理" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 清理失败" -ForegroundColor Yellow
        return $false
    }
}

function Reset-NetBIOSSessions {
    Write-Host "  → 正在刷新 NetBIOS 会话..." -ForegroundColor Cyan
    try {
        $null = nbtstat -RR 2>&1
        Write-Host "    ✓ NetBIOS 名称已重新注册" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ⚠ 操作失败" -ForegroundColor Yellow
        return $false
    }
}

function Optimize-TCPSettings {
    Write-Host "  → 正在优化 TCP 设置..." -ForegroundColor Cyan
    try {
        $null = netsh int tcp set global autotuninglevel=normal 2>&1
        $null = netsh int tcp set global rss=enabled 2>&1
        Write-Host "    ✓ TCP 设置已优化" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ⚠ 部分优化失败" -ForegroundColor Yellow
        return $false
    }
}

function Register-DNS {
    Write-Host "  → 正在重新注册 DNS..." -ForegroundColor Cyan
    try {
        $null = ipconfig /registerdns 2>&1
        Write-Host "    ✓ DNS 已重新注册" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "    ✗ 注册失败" -ForegroundColor Red
        return $false
    }
}

function Invoke-QuickOptimize {
    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host "  【快速优化】" -ForegroundColor Cyan
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ""

    $success = 0
    if (Clear-DNSCache) { $success++ }
    if (Clear-ARPCache) { $success++ }
    if (Clear-NetBIOSCache) { $success++ }

    Write-Host ""
    Write-Host "  快速优化完成！成功 $success/3 项" -ForegroundColor Green
}

function Invoke-FullOptimize {
    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host "  【完整优化】" -ForegroundColor Cyan
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ""

    $success = 0
    if (Clear-DNSCache) { $success++ }
    if (Clear-ARPCache) { $success++ }
    if (Clear-NetBIOSCache) { $success++ }
    if (Reset-NetBIOSSessions) { $success++ }
    if (Optimize-TCPSettings) { $success++ }
    if (Register-DNS) { $success++ }

    Write-Host ""
    Write-Host "  完整优化完成！成功 $success/6 项" -ForegroundColor Green
}

function Invoke-PerformanceComparison {
    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host "  【性能对比测试】" -ForegroundColor Cyan
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ""

    Write-Host "  测试优化前性能..." -ForegroundColor Yellow
    $before = Get-NetworkStats

    Write-Host ""
    Write-Host "  执行优化..." -ForegroundColor Yellow
    Invoke-QuickOptimize

    Write-Host ""
    Write-Host "  等待系统稳定..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    Write-Host ""
    Write-Host "  测试优化后性能..." -ForegroundColor Yellow
    $after = Get-NetworkStats

    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host "  性能对比结果:" -ForegroundColor White
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ""

    function Compare-Metric($Name, $Before, $After, $Unit) {
        if ($Before -lt 0 -or $After -lt 0) { return }

        $diff = $Before - $After
        $percent = if ($Before -gt 0) { ($diff / $Before) * 100 } else { 0 }
        $color = if ($diff -gt 0) { "Green" } elseif ($diff -lt 0) { "Red" } else { "White" }
        $arrow = if ($diff -gt 0) { "↓" } elseif ($diff -lt 0) { "↑" } else { "→" }

        Write-Host ("    {0}:" -f $Name)
        Write-Host ("      优化前: {0:N0} {1}" -f $Before, $Unit)
        Write-Host ("      优化后: {0:N0} {1}" -f $After, $Unit)
        Write-Host ("      变化:   " + $arrow + " " + [math]::Abs($percent).ToString("N1") + "%") -ForegroundColor $color
    }

    Compare-Metric "DNS 解析时间" $before.DNSResolveTime $after.DNSResolveTime "ms"
    Compare-Metric "Ping 延迟" $before.PingTime $after.PingTime "ms"
    Compare-Metric "HTTP 响应时间" $before.HTTPResponseTime $after.HTTPResponseTime "ms"
    Compare-Metric "DNS 缓存条目" $before.DNSCacheEntries $after.DNSCacheEntries "条"
    Compare-Metric "ARP 缓存条目" $before.ARPCacheEntries $after.ARPCacheEntries "条"
}

function Show-DetailedInfo {
    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host "  【详细网络信息】" -ForegroundColor Cyan
    Write-Host ("─" * 65) -ForegroundColor Blue
    Write-Host ""

    # 网络适配器
    Write-Host "  活动的网络适配器:" -ForegroundColor White
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 3
    foreach ($a in $adapters) {
        Write-Host ("    → {0} ({1})" -f $a.Name, $a.InterfaceDescription)
    }

    # TCP 连接统计
    Write-Host ""
    Write-Host "  TCP 连接状态统计:" -ForegroundColor White
    $tcpStats = Get-NetTCPConnection | Group-Object State | Sort-Object Count -Descending | Select-Object -First 5
    foreach ($s in $tcpStats) {
        Write-Host ("    {0}: {1}" -f $s.Name, $s.Count)
    }

    # 默认网关
    Write-Host ""
    Write-Host "  默认网关:" -ForegroundColor White
    $gateways = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Select-Object -First 2
    foreach ($g in $gateways) {
        Write-Host ("    → " + $g.NextHop)
    }
}

function Start-RealtimeMonitor {
    Write-Host ""
    Write-Host "  实时监控已启动 (按 Ctrl+C 停止)" -ForegroundColor Green
    Write-Host ""

    try {
        while ($true) {
            Clear-Host
            Write-Host ""
            Write-Host ("=" * 65) -ForegroundColor Cyan
            Write-Host "      Windows 网络监控与优化工具 (Network Monitor) v1.0" -ForegroundColor White
            Write-Host ("=" * 65) -ForegroundColor Cyan
            Write-Host "  [实时监控模式 - 按 Ctrl+C 退出]" -ForegroundColor DarkGray

            $stats = Get-NetworkStats
            Show-NetworkStats -Stats $stats

            Write-Host ""
            Write-Host "  下次刷新: 30 秒后..." -ForegroundColor DarkGray
            Start-Sleep -Seconds 30
        }
    }
    catch {
        Write-Host ""
        Write-Host "  监控已停止" -ForegroundColor Yellow
    }
}

# ============== 主程序 ==============

function Show-Menu {
    Write-Host ""
    Write-Host ("─" * 65) -ForegroundColor Cyan
    Write-Host "请选择操作:" -ForegroundColor White
    Write-Host "  1. 检测当前网络状态"
    Write-Host "  2. 快速优化 (清理 DNS/ARP/NetBIOS 缓存)"
    Write-Host "  3. 完整优化 (包含 TCP 优化)"
    Write-Host "  4. 启动实时监控"
    Write-Host "  5. 查看详细网络信息"
    Write-Host "  6. 性能对比测试"
    Write-Host "  0. 退出"
    Write-Host ("─" * 65) -ForegroundColor Cyan

    return Read-Host "`n请输入选项 [0-6]"
}

# 主循环
Clear-Host
Write-Host ""
Write-Host ("=" * 65) -ForegroundColor Cyan
Write-Host "      Windows 网络监控与优化工具 (Network Monitor) v1.0" -ForegroundColor White
Write-Host ("=" * 65) -ForegroundColor Cyan
Write-Host "  解决长时间运行后网页变慢的问题 | 实时监控网络状态" -ForegroundColor Yellow

while ($true) {
    $choice = Show-Menu

    switch ($choice) {
        '0' {
            Write-Host "`n感谢使用网络监控与优化工具！`n" -ForegroundColor Cyan
            exit
        }
        '1' {
            $stats = Get-NetworkStats
            Show-NetworkStats -Stats $stats
        }
        '2' {
            Invoke-QuickOptimize
        }
        '3' {
            Invoke-FullOptimize
        }
        '4' {
            Start-RealtimeMonitor
        }
        '5' {
            Show-DetailedInfo
        }
        '6' {
            Invoke-PerformanceComparison
        }
        default {
            Write-Host "`n无效选项，请重新选择" -ForegroundColor Red
        }
    }
}
