#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 网络性能智能优化工具 (Smart Network Optimizer)
通过单项测试定位性能问题的具体原因

作者: Network Doctor Team
版本: 2.0.0
"""

import subprocess
import sys
import os
import ctypes
import time
import socket
import urllib.request
from datetime import datetime
from typing import Tuple, Dict, List
from dataclasses import dataclass


class Colors:
    RESET = '\033[0m'
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    DIM = '\033[2m'


def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_command(command: str, timeout: int = 30) -> Tuple[int, str, str]:
    try:
        result = subprocess.run(
            command, shell=True, capture_output=True, text=True,
            timeout=timeout, encoding='gbk', errors='ignore'
        )
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "超时"
    except Exception as e:
        return -1, "", str(e)


def print_header():
    print(f"\n{Colors.CYAN}{'='*65}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.WHITE}      Windows 网络性能智能优化工具 v2.0{Colors.RESET}")
    print(f"{Colors.CYAN}{'='*65}{Colors.RESET}")
    print(f"{Colors.YELLOW}  解决长时间运行后网页变慢的问题 | 单项测试定位原因{Colors.RESET}\n")


def print_section(title: str):
    print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
    print(f"  {Colors.BOLD}{Colors.CYAN}【{title}】{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}")


# ============== 性能测试函数 ==============

@dataclass
class PerformanceResult:
    """性能测试结果"""
    dns_time: float  # DNS 解析时间 (ms)
    http_time: float  # HTTP 响应时间 (ms)
    ping_time: float  # Ping 延迟 (ms)


def measure_dns_time(domain: str = "www.baidu.com") -> float:
    """测量 DNS 解析时间"""
    try:
        start = time.perf_counter()
        socket.gethostbyname(domain)
        return (time.perf_counter() - start) * 1000
    except:
        return -1


def measure_http_time(url: str = "http://www.baidu.com") -> float:
    """测量 HTTP 响应时间"""
    try:
        start = time.perf_counter()
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            _ = resp.read(1024)
        return (time.perf_counter() - start) * 1000
    except:
        return -1


def measure_ping_time(host: str = "114.114.114.114") -> float:
    """测量 Ping 延迟"""
    ret, stdout, _ = run_command(f"ping -n 1 -w 3000 {host}", timeout=5)
    if ret == 0:
        for part in stdout.split():
            if "时间=" in part or "time=" in part.lower():
                try:
                    return float(part.split('=')[1].replace('ms', '').replace('毫秒', ''))
                except:
                    pass
    return -1


def run_performance_test(runs: int = 3) -> PerformanceResult:
    """运行性能测试（多次取平均值）"""
    dns_times = []
    http_times = []
    ping_times = []

    for i in range(runs):
        if i > 0:
            time.sleep(0.5)

        dns = measure_dns_time()
        if dns > 0:
            dns_times.append(dns)

        http = measure_http_time()
        if http > 0:
            http_times.append(http)

        ping = measure_ping_time()
        if ping > 0:
            ping_times.append(ping)

    return PerformanceResult(
        dns_time=sum(dns_times) / len(dns_times) if dns_times else -1,
        http_time=sum(http_times) / len(http_times) if http_times else -1,
        ping_time=sum(ping_times) / len(ping_times) if ping_times else -1
    )


def print_performance(result: PerformanceResult, label: str = ""):
    """打印性能结果"""
    if label:
        print(f"\n  {Colors.BOLD}{label}{Colors.RESET}")

    def get_color(value, good, warn):
        if value < 0:
            return Colors.RED
        elif value < good:
            return Colors.GREEN
        elif value < warn:
            return Colors.YELLOW
        else:
            return Colors.RED

    dns_color = get_color(result.dns_time, 50, 200)
    http_color = get_color(result.http_time, 500, 1500)
    ping_color = get_color(result.ping_time, 30, 100)

    print(f"    DNS 解析:  {dns_color}{result.dns_time:>8.1f} ms{Colors.RESET}")
    print(f"    HTTP 响应: {http_color}{result.http_time:>8.1f} ms{Colors.RESET}")
    print(f"    Ping 延迟: {ping_color}{result.ping_time:>8.1f} ms{Colors.RESET}")


# ============== 缓存状态检测 ==============

@dataclass
class CacheStatus:
    """缓存状态"""
    name: str
    count: int
    threshold: int  # 警告阈值
    description: str
    clear_function: callable


def get_dns_cache_count() -> int:
    """获取 DNS 缓存条目数"""
    ret, stdout, _ = run_command("ipconfig /displaydns", timeout=10)
    if ret == 0:
        return stdout.count("记录名称") + stdout.count("Record Name")
    return -1


def get_arp_cache_count() -> int:
    """获取 ARP 缓存条目数"""
    ret, stdout, _ = run_command("arp -a", timeout=5)
    if ret == 0:
        lines = [l for l in stdout.split('\n') if '.' in l and '-' in l]
        return len(lines)
    return -1


def get_netbios_cache_count() -> int:
    """获取 NetBIOS 缓存条目数"""
    ret, stdout, _ = run_command("nbtstat -c", timeout=5)
    if ret == 0:
        # 统计有效条目
        lines = [l for l in stdout.split('\n') if '<' in l and '>' in l]
        return len(lines)
    return 0


def get_tcp_connection_count() -> int:
    """获取 TCP 连接数"""
    ret, stdout, _ = run_command("netstat -an | find /c \"TCP\"", timeout=15)
    if ret == 0:
        try:
            return int(stdout.strip())
        except:
            pass
    return -1


# ============== 清理函数 ==============

def clear_dns_cache() -> Tuple[bool, str]:
    """清理 DNS 缓存"""
    ret, _, stderr = run_command("ipconfig /flushdns")
    return ret == 0, "DNS 缓存已清理" if ret == 0 else f"失败: {stderr}"


def clear_arp_cache() -> Tuple[bool, str]:
    """清理 ARP 缓存"""
    run_command("netsh interface ip delete arpcache")
    ret, _, _ = run_command("arp -d *")
    return True, "ARP 缓存已清理"


def clear_netbios_cache() -> Tuple[bool, str]:
    """清理 NetBIOS 缓存"""
    ret, _, _ = run_command("nbtstat -R")
    return ret == 0, "NetBIOS 缓存已清理" if ret == 0 else "清理失败"


def refresh_netbios() -> Tuple[bool, str]:
    """刷新 NetBIOS 注册"""
    ret, _, _ = run_command("nbtstat -RR")
    return ret == 0, "NetBIOS 已重新注册" if ret == 0 else "刷新失败"


# ============== 智能优化流程 ==============

def show_cache_status():
    """显示当前缓存状态"""
    print_section("当前缓存状态")

    caches = [
        ("DNS 缓存", get_dns_cache_count(), 500, "条目过多会降低 DNS 查询效率"),
        ("ARP 缓存", get_arp_cache_count(), 100, "条目过多会增加地址解析延迟"),
        ("NetBIOS 缓存", get_netbios_cache_count(), 50, "条目过多可能影响网络名称解析"),
        ("TCP 连接数", get_tcp_connection_count(), 300, "连接过多可能占用系统资源"),
    ]

    for name, count, threshold, desc in caches:
        if count < 0:
            status = f"{Colors.YELLOW}[未知]{Colors.RESET}"
        elif count > threshold:
            status = f"{Colors.RED}[偏高]{Colors.RESET}"
        else:
            status = f"{Colors.GREEN}[正常]{Colors.RESET}"

        print(f"\n  {status} {name}: {count if count >= 0 else '无法获取'}")
        if count > threshold:
            print(f"       {Colors.DIM}({desc}){Colors.RESET}")

    return caches


def single_optimize_test():
    """单项优化测试：测试每个优化操作的效果"""
    print_section("单项优化效果测试")

    optimizations = [
        ("清理 DNS 缓存", "影响域名解析速度", clear_dns_cache),
        ("清理 ARP 缓存", "影响网络地址解析", clear_arp_cache),
        ("清理 NetBIOS 缓存", "影响 Windows 网络名称解析", clear_netbios_cache),
        ("刷新 NetBIOS 注册", "重新注册网络名称", refresh_netbios),
    ]

    print("\n  可测试的优化项：")
    for i, (name, desc, _) in enumerate(optimizations, 1):
        print(f"  {Colors.WHITE}{i}{Colors.RESET}. {name}")
        print(f"     {Colors.DIM}{desc}{Colors.RESET}")
    print(f"  {Colors.WHITE}0{Colors.RESET}. 返回")

    choice = input(f"\n  请选择 [0-{len(optimizations)}]: ").strip()

    try:
        idx = int(choice)
        if idx == 0:
            return
        if idx < 1 or idx > len(optimizations):
            print(f"  {Colors.RED}无效选项{Colors.RESET}")
            return
    except:
        print(f"  {Colors.RED}无效输入{Colors.RESET}")
        return

    name, desc, func = optimizations[idx - 1]

    # 优化前测试
    print(f"\n  {Colors.CYAN}测试优化前性能...{Colors.RESET}")
    before = run_performance_test(runs=3)
    print_performance(before, "优化前:")

    # 执行优化
    print(f"\n  {Colors.CYAN}执行: {name}...{Colors.RESET}")
    success, message = func()
    print(f"  {'✓' if success else '✗'} {message}")

    if not success:
        return

    # 等待生效
    print(f"\n  {Colors.DIM}等待系统稳定...{Colors.RESET}")
    time.sleep(2)

    # 优化后测试
    print(f"\n  {Colors.CYAN}测试优化后性能...{Colors.RESET}")
    after = run_performance_test(runs=3)
    print_performance(after, "优化后:")

    # 效果分析
    print(f"\n  {Colors.BOLD}效果分析:{Colors.RESET}")

    def analyze(name, before_val, after_val):
        if before_val < 0 or after_val < 0:
            return

        diff = before_val - after_val
        percent = (diff / before_val * 100) if before_val > 0 else 0

        if abs(percent) < 5:
            print(f"    {name}: 无明显变化")
        elif diff > 0:
            print(f"    {name}: {Colors.GREEN}提升 {percent:.1f}% ({diff:.0f}ms){Colors.RESET}")
        else:
            print(f"    {name}: {Colors.RED}下降 {-percent:.1f}% ({-diff:.0f}ms){Colors.RESET}")

    analyze("DNS 解析", before.dns_time, after.dns_time)
    analyze("HTTP 响应", before.http_time, after.http_time)
    analyze("Ping 延迟", before.ping_time, after.ping_time)

    # 结论
    http_improved = before.http_time > 0 and after.http_time > 0 and after.http_time < before.http_time * 0.9
    dns_improved = before.dns_time > 0 and after.dns_time > 0 and after.dns_time < before.dns_time * 0.9

    if http_improved or dns_improved:
        print(f"\n  {Colors.GREEN}{'─'*50}{Colors.RESET}")
        print(f"  {Colors.GREEN}结论: {name} 对性能有明显改善作用{Colors.RESET}")
        print(f"  {Colors.GREEN}建议: 定期执行此优化操作{Colors.RESET}")
        print(f"  {Colors.GREEN}{'─'*50}{Colors.RESET}")

        # 保存结果
        save_optimization_result(name, before, after)
    else:
        print(f"\n  {Colors.YELLOW}结论: {name} 对当前性能无明显影响{Colors.RESET}")


def comprehensive_test():
    """综合测试：逐一测试所有优化项，找出最有效的"""
    print_section("综合效果测试")

    print(f"\n  {Colors.YELLOW}将逐一测试每个优化项的效果，这可能需要几分钟时间{Colors.RESET}")
    confirm = input("  继续? (y/n): ").strip().lower()
    if confirm != 'y':
        return

    # 基准测试
    print(f"\n  {Colors.CYAN}[1/5] 基准性能测试...{Colors.RESET}")
    baseline = run_performance_test(runs=5)
    print_performance(baseline, "基准性能:")

    optimizations = [
        ("DNS 缓存清理", clear_dns_cache),
        ("ARP 缓存清理", clear_arp_cache),
        ("NetBIOS 缓存清理", clear_netbios_cache),
        ("NetBIOS 重新注册", refresh_netbios),
    ]

    results = []

    for i, (name, func) in enumerate(optimizations, 2):
        print(f"\n  {Colors.CYAN}[{i}/5] 测试: {name}...{Colors.RESET}")

        # 执行优化
        success, _ = func()
        if not success:
            print(f"  {Colors.YELLOW}跳过 (执行失败){Colors.RESET}")
            continue

        time.sleep(1)

        # 测试效果
        perf = run_performance_test(runs=3)

        # 计算改善
        http_improve = 0
        if baseline.http_time > 0 and perf.http_time > 0:
            http_improve = (baseline.http_time - perf.http_time) / baseline.http_time * 100

        dns_improve = 0
        if baseline.dns_time > 0 and perf.dns_time > 0:
            dns_improve = (baseline.dns_time - perf.dns_time) / baseline.dns_time * 100

        results.append({
            'name': name,
            'http_time': perf.http_time,
            'http_improve': http_improve,
            'dns_time': perf.dns_time,
            'dns_improve': dns_improve,
        })

        status = f"{Colors.GREEN}有效{Colors.RESET}" if http_improve > 5 or dns_improve > 5 else f"{Colors.DIM}效果不明显{Colors.RESET}"
        print(f"  HTTP: {perf.http_time:.0f}ms (变化: {http_improve:+.1f}%) - {status}")

    # 汇总结果
    print_section("测试结果汇总")

    print(f"\n  {Colors.BOLD}各优化项效果排名 (按 HTTP 响应时间改善):{Colors.RESET}\n")

    sorted_results = sorted(results, key=lambda x: x['http_improve'], reverse=True)

    for i, r in enumerate(sorted_results, 1):
        if r['http_improve'] > 5:
            color = Colors.GREEN
            verdict = "有效"
        elif r['http_improve'] > 0:
            color = Colors.YELLOW
            verdict = "轻微"
        else:
            color = Colors.DIM
            verdict = "无效"

        print(f"  {i}. {r['name']}")
        print(f"     HTTP 改善: {color}{r['http_improve']:+.1f}%{Colors.RESET} [{verdict}]")
        print(f"     DNS 改善:  {r['dns_improve']:+.1f}%")

    # 结论
    effective = [r for r in results if r['http_improve'] > 5]
    if effective:
        print(f"\n  {Colors.GREEN}{'─'*50}{Colors.RESET}")
        print(f"  {Colors.GREEN}建议定期执行以下优化:{Colors.RESET}")
        for r in effective:
            print(f"    → {r['name']} (改善 {r['http_improve']:.1f}%)")
        print(f"  {Colors.GREEN}{'─'*50}{Colors.RESET}")
    else:
        print(f"\n  {Colors.YELLOW}当前网络性能良好，无需特别优化{Colors.RESET}")


def save_optimization_result(name: str, before: PerformanceResult, after: PerformanceResult):
    """保存优化结果"""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("optimization_log.txt", "a", encoding="utf-8") as f:
            f.write(f"\n[{timestamp}]\n")
            f.write(f"优化操作: {name}\n")
            f.write(f"HTTP 响应: {before.http_time:.0f}ms → {after.http_time:.0f}ms\n")
            f.write(f"DNS 解析: {before.dns_time:.0f}ms → {after.dns_time:.0f}ms\n")
            f.write("-" * 40 + "\n")
        print(f"\n  {Colors.DIM}结果已保存到 optimization_log.txt{Colors.RESET}")
    except:
        pass


def show_menu():
    print(f"\n{Colors.CYAN}{'─'*65}{Colors.RESET}")
    print(f"{Colors.BOLD}请选择操作:{Colors.RESET}")
    print(f"  {Colors.WHITE}1{Colors.RESET}. 查看当前缓存状态")
    print(f"  {Colors.WHITE}2{Colors.RESET}. 测试当前网络性能")
    print(f"  {Colors.WHITE}3{Colors.RESET}. 单项优化测试 (测试单个优化的效果)")
    print(f"  {Colors.WHITE}4{Colors.RESET}. 综合效果测试 (测试所有优化项，找出最有效的)")
    print(f"  {Colors.WHITE}0{Colors.RESET}. 退出")
    print(f"{Colors.CYAN}{'─'*65}{Colors.RESET}")
    return input(f"\n{Colors.BOLD}请输入选项 [0-4]: {Colors.RESET}").strip()


def main():
    if sys.platform == 'win32':
        os.system('color')
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
        except:
            pass

    print_header()

    if not is_admin():
        print(f"{Colors.YELLOW}警告: 程序未以管理员权限运行，部分功能可能无法使用{Colors.RESET}")

    while True:
        choice = show_menu()

        if choice == '0':
            print(f"\n{Colors.CYAN}感谢使用！{Colors.RESET}\n")
            break
        elif choice == '1':
            show_cache_status()
        elif choice == '2':
            print_section("网络性能测试")
            print(f"\n  {Colors.CYAN}正在测试 (测量 3 次取平均值)...{Colors.RESET}")
            result = run_performance_test(runs=3)
            print_performance(result, "当前性能:")
        elif choice == '3':
            single_optimize_test()
        elif choice == '4':
            comprehensive_test()
        else:
            print(f"\n{Colors.RED}无效选项{Colors.RESET}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}程序被中断{Colors.RESET}")
    except Exception as e:
        print(f"\n{Colors.RED}程序出错: {e}{Colors.RESET}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")
