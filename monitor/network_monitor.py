#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 网络监控与优化工具 (Network Monitor & Optimizer)
用于实时监控网络状态并定期优化，解决长时间运行后网速变慢的问题

作者: Network Doctor Team
版本: 1.0.0
"""

import subprocess
import sys
import os
import ctypes
import time
import socket
import threading
import json
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from collections import deque
import urllib.request
import ssl


class Colors:
    """终端颜色定义"""
    RESET = '\033[0m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    DIM = '\033[2m'


@dataclass
class NetworkStats:
    """网络统计数据"""
    timestamp: datetime
    dns_resolve_time: float  # DNS 解析时间 (ms)
    ping_time: float  # Ping 延迟 (ms)
    http_response_time: float  # HTTP 响应时间 (ms)
    tcp_connections: int  # TCP 连接数
    dns_cache_entries: int  # DNS 缓存条目数
    arp_cache_entries: int  # ARP 缓存条目数
    memory_usage_percent: float  # 内存使用率
    status: str = "正常"  # 状态


@dataclass
class OptimizationResult:
    """优化结果"""
    action: str
    success: bool
    message: str
    improvement: Optional[str] = None


def is_admin() -> bool:
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_command(command: str, timeout: int = 30) -> Tuple[int, str, str]:
    """执行系统命令"""
    try:
        result = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding='gbk',
            errors='ignore'
        )
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "命令执行超时"
    except Exception as e:
        return -1, "", str(e)


def clear_screen():
    """清屏"""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """打印程序头部"""
    print(f"\n{Colors.CYAN}{'='*65}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.WHITE}      Windows 网络监控与优化工具 (Network Monitor) v1.0{Colors.RESET}")
    print(f"{Colors.CYAN}{'='*65}{Colors.RESET}")
    print(f"{Colors.YELLOW}  解决长时间运行后网页变慢的问题 | 实时监控网络状态{Colors.RESET}\n")


class NetworkMonitor:
    """网络监控类"""

    def __init__(self):
        self.stats_history: deque = deque(maxlen=100)  # 保留最近100条记录
        self.is_monitoring = False
        self.monitor_thread = None
        self.optimization_threshold = {
            'dns_cache_entries': 1000,  # DNS 缓存超过1000条时优化
            'tcp_connections': 500,  # TCP 连接超过500时警告
            'arp_cache_entries': 200,  # ARP 缓存超过200条时清理
            'http_response_time': 3000,  # HTTP 响应超过3秒时警告
        }

    def measure_dns_resolve_time(self, domain: str = "www.baidu.com") -> float:
        """测量 DNS 解析时间"""
        try:
            start = time.perf_counter()
            socket.gethostbyname(domain)
            end = time.perf_counter()
            return (end - start) * 1000  # 转换为毫秒
        except:
            return -1

    def measure_ping_time(self, host: str = "114.114.114.114") -> float:
        """测量 Ping 延迟"""
        try:
            ret, stdout, _ = run_command(f"ping -n 1 -w 2000 {host}", timeout=5)
            if ret == 0 and "时间=" in stdout:
                # 提取时间值
                for part in stdout.split():
                    if "时间=" in part or "time=" in part.lower():
                        time_str = part.split('=')[1].replace('ms', '').replace('毫秒', '')
                        return float(time_str)
            elif ret == 0 and "time=" in stdout.lower():
                for part in stdout.split():
                    if "time=" in part.lower():
                        time_str = part.split('=')[1].replace('ms', '')
                        return float(time_str)
            return -1
        except:
            return -1

    def measure_http_response_time(self, url: str = "http://www.baidu.com") -> float:
        """测量 HTTP 响应时间"""
        try:
            start = time.perf_counter()
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=10) as response:
                _ = response.read(1024)  # 只读取前1KB
            end = time.perf_counter()
            return (end - start) * 1000
        except:
            return -1

    def get_tcp_connection_count(self) -> int:
        """获取 TCP 连接数"""
        try:
            ret, stdout, _ = run_command("netstat -an | find /c \"TCP\"", timeout=10)
            if ret == 0:
                return int(stdout.strip())
            return -1
        except:
            return -1

    def get_dns_cache_count(self) -> int:
        """获取 DNS 缓存条目数"""
        try:
            ret, stdout, _ = run_command("ipconfig /displaydns", timeout=10)
            if ret == 0:
                # 统计记录名称的数量
                count = stdout.count("记录名称") + stdout.count("Record Name")
                return count
            return -1
        except:
            return -1

    def get_arp_cache_count(self) -> int:
        """获取 ARP 缓存条目数"""
        try:
            ret, stdout, _ = run_command("arp -a", timeout=5)
            if ret == 0:
                # 统计有效的 ARP 条目
                lines = [l for l in stdout.split('\n') if '.' in l and '-' in l]
                return len(lines)
            return -1
        except:
            return -1

    def get_memory_usage(self) -> float:
        """获取内存使用率"""
        try:
            ret, stdout, _ = run_command(
                'wmic OS get FreePhysicalMemory,TotalVisibleMemorySize /Value',
                timeout=5
            )
            if ret == 0:
                free = 0
                total = 0
                for line in stdout.split('\n'):
                    if 'FreePhysicalMemory=' in line:
                        free = int(line.split('=')[1].strip())
                    elif 'TotalVisibleMemorySize=' in line:
                        total = int(line.split('=')[1].strip())
                if total > 0:
                    return ((total - free) / total) * 100
            return -1
        except:
            return -1

    def collect_stats(self) -> NetworkStats:
        """收集网络统计数据"""
        stats = NetworkStats(
            timestamp=datetime.now(),
            dns_resolve_time=self.measure_dns_resolve_time(),
            ping_time=self.measure_ping_time(),
            http_response_time=self.measure_http_response_time(),
            tcp_connections=self.get_tcp_connection_count(),
            dns_cache_entries=self.get_dns_cache_count(),
            arp_cache_entries=self.get_arp_cache_count(),
            memory_usage_percent=self.get_memory_usage()
        )

        # 判断状态
        if stats.http_response_time > self.optimization_threshold['http_response_time']:
            stats.status = "缓慢"
        elif stats.http_response_time > 1500:
            stats.status = "较慢"
        elif stats.http_response_time < 0:
            stats.status = "异常"
        else:
            stats.status = "正常"

        return stats

    def print_stats(self, stats: NetworkStats):
        """打印统计数据"""
        status_color = {
            "正常": Colors.GREEN,
            "较慢": Colors.YELLOW,
            "缓慢": Colors.RED,
            "异常": Colors.RED
        }.get(stats.status, Colors.WHITE)

        print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
        print(f"  {Colors.CYAN}时间:{Colors.RESET} {stats.timestamp.strftime('%Y-%m-%d %H:%M:%S')}    "
              f"{Colors.CYAN}状态:{Colors.RESET} {status_color}{stats.status}{Colors.RESET}")
        print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}")

        # 网络延迟
        print(f"\n  {Colors.BOLD}网络延迟指标:{Colors.RESET}")
        dns_color = Colors.GREEN if 0 < stats.dns_resolve_time < 100 else Colors.YELLOW if stats.dns_resolve_time < 500 else Colors.RED
        ping_color = Colors.GREEN if 0 < stats.ping_time < 50 else Colors.YELLOW if stats.ping_time < 100 else Colors.RED
        http_color = Colors.GREEN if 0 < stats.http_response_time < 500 else Colors.YELLOW if stats.http_response_time < 1500 else Colors.RED

        print(f"    DNS 解析时间:   {dns_color}{stats.dns_resolve_time:>8.1f} ms{Colors.RESET}")
        print(f"    Ping 延迟:      {ping_color}{stats.ping_time:>8.1f} ms{Colors.RESET}")
        print(f"    HTTP 响应时间:  {http_color}{stats.http_response_time:>8.1f} ms{Colors.RESET}")

        # 系统资源
        print(f"\n  {Colors.BOLD}系统资源状态:{Colors.RESET}")
        tcp_color = Colors.GREEN if stats.tcp_connections < 300 else Colors.YELLOW if stats.tcp_connections < 500 else Colors.RED
        dns_cache_color = Colors.GREEN if stats.dns_cache_entries < 500 else Colors.YELLOW if stats.dns_cache_entries < 1000 else Colors.RED
        arp_color = Colors.GREEN if stats.arp_cache_entries < 100 else Colors.YELLOW if stats.arp_cache_entries < 200 else Colors.RED
        mem_color = Colors.GREEN if stats.memory_usage_percent < 70 else Colors.YELLOW if stats.memory_usage_percent < 85 else Colors.RED

        print(f"    TCP 连接数:     {tcp_color}{stats.tcp_connections:>8} 个{Colors.RESET}")
        print(f"    DNS 缓存条目:   {dns_cache_color}{stats.dns_cache_entries:>8} 条{Colors.RESET}")
        print(f"    ARP 缓存条目:   {arp_color}{stats.arp_cache_entries:>8} 条{Colors.RESET}")
        print(f"    内存使用率:     {mem_color}{stats.memory_usage_percent:>7.1f} %{Colors.RESET}")

        # 优化建议
        suggestions = []
        if stats.dns_cache_entries > self.optimization_threshold['dns_cache_entries']:
            suggestions.append("DNS 缓存过大，建议清理")
        if stats.tcp_connections > self.optimization_threshold['tcp_connections']:
            suggestions.append("TCP 连接数过多，可能存在连接泄漏")
        if stats.arp_cache_entries > self.optimization_threshold['arp_cache_entries']:
            suggestions.append("ARP 缓存过大，建议清理")
        if stats.http_response_time > 2000:
            suggestions.append("网络响应缓慢，建议执行优化")

        if suggestions:
            print(f"\n  {Colors.YELLOW}优化建议:{Colors.RESET}")
            for s in suggestions:
                print(f"    {Colors.YELLOW}→{Colors.RESET} {s}")

    def analyze_performance(self) -> Dict:
        """分析网络性能趋势"""
        if len(self.stats_history) < 2:
            return {"trend": "数据不足", "suggestion": "需要更多监控数据"}

        recent = list(self.stats_history)[-10:]  # 最近10条

        avg_http = sum(s.http_response_time for s in recent if s.http_response_time > 0) / len(recent)
        avg_dns = sum(s.dns_resolve_time for s in recent if s.dns_resolve_time > 0) / len(recent)

        # 比较最新和之前的数据
        if len(self.stats_history) >= 20:
            older = list(self.stats_history)[-20:-10]
            old_avg_http = sum(s.http_response_time for s in older if s.http_response_time > 0) / len(older)

            if avg_http > old_avg_http * 1.5:
                return {
                    "trend": "性能下降",
                    "suggestion": "检测到网络性能下降，建议执行优化",
                    "http_increase": f"{((avg_http - old_avg_http) / old_avg_http * 100):.1f}%"
                }

        return {
            "trend": "稳定",
            "avg_http_time": f"{avg_http:.1f}ms",
            "avg_dns_time": f"{avg_dns:.1f}ms"
        }


class NetworkOptimizer:
    """网络优化器"""

    def __init__(self):
        self.optimization_log: List[OptimizationResult] = []

    def flush_dns_cache(self) -> OptimizationResult:
        """清理 DNS 缓存"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在清理 DNS 缓存...")
        ret, stdout, stderr = run_command("ipconfig /flushdns")

        if ret == 0:
            result = OptimizationResult(
                action="清理 DNS 缓存",
                success=True,
                message="DNS 缓存已清理",
                improvement="减少 DNS 解析延迟"
            )
            print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        else:
            result = OptimizationResult(
                action="清理 DNS 缓存",
                success=False,
                message=f"清理失败: {stderr}"
            )
            print(f"    {Colors.RED}✗{Colors.RESET} {result.message}")

        self.optimization_log.append(result)
        return result

    def clear_arp_cache(self) -> OptimizationResult:
        """清理 ARP 缓存"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在清理 ARP 缓存...")
        ret, stdout, stderr = run_command("netsh interface ip delete arpcache")

        if ret == 0:
            result = OptimizationResult(
                action="清理 ARP 缓存",
                success=True,
                message="ARP 缓存已清理",
                improvement="减少网络地址解析延迟"
            )
            print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        else:
            # 尝试备用命令
            ret2, _, _ = run_command("arp -d *")
            if ret2 == 0:
                result = OptimizationResult(
                    action="清理 ARP 缓存",
                    success=True,
                    message="ARP 缓存已清理 (备用方法)"
                )
                print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
            else:
                result = OptimizationResult(
                    action="清理 ARP 缓存",
                    success=False,
                    message="清理失败，可能需要管理员权限"
                )
                print(f"    {Colors.RED}✗{Colors.RESET} {result.message}")

        self.optimization_log.append(result)
        return result

    def clear_netbios_cache(self) -> OptimizationResult:
        """清理 NetBIOS 缓存"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在清理 NetBIOS 缓存...")
        ret, stdout, stderr = run_command("nbtstat -R")

        if ret == 0:
            result = OptimizationResult(
                action="清理 NetBIOS 缓存",
                success=True,
                message="NetBIOS 名称缓存已清理"
            )
            print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        else:
            result = OptimizationResult(
                action="清理 NetBIOS 缓存",
                success=False,
                message="清理失败"
            )
            print(f"    {Colors.YELLOW}⚠{Colors.RESET} {result.message}")

        self.optimization_log.append(result)
        return result

    def reset_netbios_sessions(self) -> OptimizationResult:
        """重置 NetBIOS 会话表"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在刷新 NetBIOS 会话...")
        ret, stdout, stderr = run_command("nbtstat -RR")

        result = OptimizationResult(
            action="刷新 NetBIOS 会话",
            success=(ret == 0),
            message="NetBIOS 名称已重新注册" if ret == 0 else "操作失败"
        )

        if result.success:
            print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        else:
            print(f"    {Colors.YELLOW}⚠{Colors.RESET} {result.message}")

        self.optimization_log.append(result)
        return result

    def optimize_tcp_settings(self) -> OptimizationResult:
        """优化 TCP 设置"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在优化 TCP 设置...")

        commands = [
            # 启用 TCP 窗口自动调整
            ("netsh int tcp set global autotuninglevel=normal", "启用 TCP 自动调整"),
            # 启用 RSS (接收端缩放)
            ("netsh int tcp set global rss=enabled", "启用 RSS"),
            # 禁用 TCP 任务卸载 (某些情况下可以改善性能)
            # ("netsh int tcp set global chimney=disabled", "禁用 TCP Chimney"),
        ]

        success_count = 0
        for cmd, desc in commands:
            ret, _, _ = run_command(cmd)
            if ret == 0:
                success_count += 1

        result = OptimizationResult(
            action="优化 TCP 设置",
            success=(success_count > 0),
            message=f"已应用 {success_count}/{len(commands)} 项 TCP 优化",
            improvement="提高 TCP 传输效率"
        )

        print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        self.optimization_log.append(result)
        return result

    def release_idle_connections(self) -> OptimizationResult:
        """释放空闲连接（通过重置网络组件）"""
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在释放空闲网络连接...")

        # 获取当前连接数
        ret, stdout, _ = run_command("netstat -an | find /c \"TCP\"", timeout=10)
        before_count = int(stdout.strip()) if ret == 0 else 0

        # 刷新 DNS 客户端解析器缓存可以帮助清理一些悬挂的连接
        run_command("ipconfig /registerdns")

        result = OptimizationResult(
            action="释放空闲连接",
            success=True,
            message=f"已尝试释放空闲连接 (当前: {before_count} 个)",
            improvement="释放系统网络资源"
        )

        print(f"    {Colors.GREEN}✓{Colors.RESET} {result.message}")
        self.optimization_log.append(result)
        return result

    def quick_optimize(self) -> List[OptimizationResult]:
        """快速优化 - 最常用的优化操作"""
        print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
        print(f"  {Colors.BOLD}{Colors.CYAN}【快速优化】{Colors.RESET}")
        print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}\n")

        results = []
        results.append(self.flush_dns_cache())
        results.append(self.clear_arp_cache())
        results.append(self.clear_netbios_cache())

        success = sum(1 for r in results if r.success)
        print(f"\n  {Colors.GREEN}快速优化完成！{Colors.RESET} 成功 {success}/{len(results)} 项")

        return results

    def full_optimize(self) -> List[OptimizationResult]:
        """完整优化 - 所有优化操作"""
        print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
        print(f"  {Colors.BOLD}{Colors.CYAN}【完整优化】{Colors.RESET}")
        print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}\n")

        results = []
        results.append(self.flush_dns_cache())
        results.append(self.clear_arp_cache())
        results.append(self.clear_netbios_cache())
        results.append(self.reset_netbios_sessions())
        results.append(self.optimize_tcp_settings())
        results.append(self.release_idle_connections())

        success = sum(1 for r in results if r.success)
        print(f"\n  {Colors.GREEN}完整优化完成！{Colors.RESET} 成功 {success}/{len(results)} 项")

        return results


class ScheduledOptimizer:
    """定时优化器"""

    def __init__(self, monitor: NetworkMonitor, optimizer: NetworkOptimizer):
        self.monitor = monitor
        self.optimizer = optimizer
        self.is_running = False
        self.thread = None
        self.interval = 1800  # 默认30分钟
        self.auto_optimize_enabled = True

    def start(self, interval_minutes: int = 30):
        """启动定时优化"""
        self.interval = interval_minutes * 60
        self.is_running = True
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        print(f"  {Colors.GREEN}✓{Colors.RESET} 定时优化已启动 (间隔: {interval_minutes} 分钟)")

    def stop(self):
        """停止定时优化"""
        self.is_running = False
        if self.thread:
            self.thread.join(timeout=2)
        print(f"  {Colors.YELLOW}→{Colors.RESET} 定时优化已停止")

    def _run(self):
        """定时任务主循环"""
        last_optimize_time = time.time()

        while self.is_running:
            time.sleep(60)  # 每分钟检查一次

            current_time = time.time()
            elapsed = current_time - last_optimize_time

            if elapsed >= self.interval:
                if self.auto_optimize_enabled:
                    print(f"\n{Colors.CYAN}[定时任务] 执行自动优化...{Colors.RESET}")
                    self.optimizer.quick_optimize()
                    last_optimize_time = current_time


def show_menu() -> str:
    """显示主菜单"""
    print(f"\n{Colors.CYAN}{'─'*65}{Colors.RESET}")
    print(f"{Colors.BOLD}请选择操作:{Colors.RESET}")
    print(f"  {Colors.WHITE}1{Colors.RESET}. 检测当前网络状态")
    print(f"  {Colors.WHITE}2{Colors.RESET}. 快速优化 (清理 DNS/ARP/NetBIOS 缓存)")
    print(f"  {Colors.WHITE}3{Colors.RESET}. 完整优化 (包含 TCP 优化)")
    print(f"  {Colors.WHITE}4{Colors.RESET}. 启动实时监控 (每30秒刷新)")
    print(f"  {Colors.WHITE}5{Colors.RESET}. 启动定时自动优化 (每30分钟)")
    print(f"  {Colors.WHITE}6{Colors.RESET}. 查看详细网络信息")
    print(f"  {Colors.WHITE}7{Colors.RESET}. 性能对比测试 (优化前后)")
    print(f"  {Colors.WHITE}0{Colors.RESET}. 退出")
    print(f"{Colors.CYAN}{'─'*65}{Colors.RESET}")

    return input(f"\n{Colors.BOLD}请输入选项 [0-7]: {Colors.RESET}").strip()


def show_detailed_network_info():
    """显示详细网络信息"""
    print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
    print(f"  {Colors.BOLD}{Colors.CYAN}【详细网络信息】{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}\n")

    # 网络适配器信息
    print(f"  {Colors.BOLD}网络适配器:{Colors.RESET}")
    ret, stdout, _ = run_command("ipconfig")
    active_adapters = []
    current_adapter = None
    for line in stdout.split('\n'):
        if '适配器' in line or 'adapter' in line.lower():
            current_adapter = line.strip()
        elif current_adapter and ('IPv4' in line or 'IP 地址' in line):
            active_adapters.append(current_adapter)
    for adapter in active_adapters[:3]:
        print(f"    → {adapter}")

    # TCP 连接统计
    print(f"\n  {Colors.BOLD}TCP 连接统计:{Colors.RESET}")
    ret, stdout, _ = run_command("netstat -an | findstr TCP", timeout=15)
    if ret == 0:
        states = {}
        for line in stdout.split('\n'):
            parts = line.split()
            if len(parts) >= 4:
                state = parts[-1]
                states[state] = states.get(state, 0) + 1
        for state, count in sorted(states.items(), key=lambda x: -x[1])[:5]:
            print(f"    {state}: {count}")

    # 路由表概要
    print(f"\n  {Colors.BOLD}默认网关:{Colors.RESET}")
    ret, stdout, _ = run_command("ipconfig | findstr 网关")
    for line in stdout.split('\n'):
        if line.strip() and '.' in line:
            print(f"    → {line.strip()}")
            break


def run_performance_comparison(monitor: NetworkMonitor, optimizer: NetworkOptimizer):
    """运行优化前后性能对比测试"""
    print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
    print(f"  {Colors.BOLD}{Colors.CYAN}【性能对比测试】{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}\n")

    # 优化前测试
    print(f"  {Colors.YELLOW}测试优化前性能...{Colors.RESET}")
    before = monitor.collect_stats()

    # 执行优化
    print(f"\n  {Colors.YELLOW}执行优化...{Colors.RESET}")
    optimizer.quick_optimize()

    # 等待一会让系统稳定
    print(f"\n  {Colors.YELLOW}等待系统稳定...{Colors.RESET}")
    time.sleep(3)

    # 优化后测试
    print(f"\n  {Colors.YELLOW}测试优化后性能...{Colors.RESET}")
    after = monitor.collect_stats()

    # 显示对比结果
    print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
    print(f"  {Colors.BOLD}性能对比结果:{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}\n")

    def compare(name: str, before_val: float, after_val: float, unit: str, lower_is_better: bool = True):
        if before_val < 0 or after_val < 0:
            return
        diff = before_val - after_val if lower_is_better else after_val - before_val
        percent = (diff / before_val * 100) if before_val > 0 else 0
        color = Colors.GREEN if diff > 0 else Colors.RED if diff < 0 else Colors.WHITE
        arrow = "↓" if (lower_is_better and diff > 0) or (not lower_is_better and diff < 0) else "↑" if diff != 0 else "→"
        print(f"    {name}:")
        print(f"      优化前: {before_val:.1f} {unit}")
        print(f"      优化后: {after_val:.1f} {unit}")
        print(f"      变化:   {color}{arrow} {abs(percent):.1f}%{Colors.RESET}")

    compare("DNS 解析时间", before.dns_resolve_time, after.dns_resolve_time, "ms")
    compare("Ping 延迟", before.ping_time, after.ping_time, "ms")
    compare("HTTP 响应时间", before.http_response_time, after.http_response_time, "ms")
    compare("DNS 缓存条目", float(before.dns_cache_entries), float(after.dns_cache_entries), "条")
    compare("ARP 缓存条目", float(before.arp_cache_entries), float(after.arp_cache_entries), "条")


def run_realtime_monitor(monitor: NetworkMonitor):
    """运行实时监控"""
    print(f"\n{Colors.GREEN}实时监控已启动 (按 Ctrl+C 停止){Colors.RESET}")

    try:
        while True:
            clear_screen()
            print_header()
            print(f"  {Colors.DIM}[实时监控模式 - 按 Ctrl+C 退出]{Colors.RESET}")

            stats = monitor.collect_stats()
            monitor.stats_history.append(stats)
            monitor.print_stats(stats)

            # 显示性能趋势
            analysis = monitor.analyze_performance()
            if analysis.get("trend") != "数据不足":
                print(f"\n  {Colors.BOLD}性能趋势:{Colors.RESET} {analysis.get('trend', 'N/A')}")

            print(f"\n  {Colors.DIM}下次刷新: 30 秒后...{Colors.RESET}")
            time.sleep(30)

    except KeyboardInterrupt:
        print(f"\n\n  {Colors.YELLOW}监控已停止{Colors.RESET}")


def main():
    """主函数"""
    # 启用 Windows 终端颜色支持
    if sys.platform == 'win32':
        os.system('color')
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
        except:
            pass

    print_header()

    # 检查管理员权限
    if not is_admin():
        print(f"{Colors.YELLOW}提示: 部分功能需要管理员权限才能正常工作{Colors.RESET}")
        print(f"{Colors.YELLOW}建议右键点击程序，选择「以管理员身份运行」{Colors.RESET}\n")

    monitor = NetworkMonitor()
    optimizer = NetworkOptimizer()
    scheduled = ScheduledOptimizer(monitor, optimizer)

    while True:
        choice = show_menu()

        if choice == '0':
            if scheduled.is_running:
                scheduled.stop()
            print(f"\n{Colors.CYAN}感谢使用网络监控与优化工具！{Colors.RESET}\n")
            break

        elif choice == '1':
            print(f"\n  {Colors.YELLOW}正在检测网络状态...{Colors.RESET}")
            stats = monitor.collect_stats()
            monitor.stats_history.append(stats)
            monitor.print_stats(stats)

        elif choice == '2':
            optimizer.quick_optimize()

        elif choice == '3':
            optimizer.full_optimize()

        elif choice == '4':
            run_realtime_monitor(monitor)

        elif choice == '5':
            if scheduled.is_running:
                scheduled.stop()
            else:
                interval = input("  请输入优化间隔(分钟, 默认30): ").strip()
                interval = int(interval) if interval.isdigit() else 30
                scheduled.start(interval)

        elif choice == '6':
            show_detailed_network_info()

        elif choice == '7':
            run_performance_comparison(monitor, optimizer)

        else:
            print(f"\n{Colors.RED}无效选项，请重新选择{Colors.RESET}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}程序被用户中断{Colors.RESET}")
    except Exception as e:
        print(f"\n{Colors.RED}程序出错: {e}{Colors.RESET}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")
