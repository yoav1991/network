#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 网络智能诊断工具 (Smart Network Diagnose)
通过逐步测试定位具体问题，而不是一次性执行所有修复

作者: Network Doctor Team
版本: 2.0.0
"""

import subprocess
import sys
import os
import ctypes
import winreg
import socket
import time
import urllib.request
import ssl
from datetime import datetime
from typing import Dict, List, Tuple, Optional
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
    print(f"{Colors.BOLD}{Colors.WHITE}      Windows 网络智能诊断工具 v2.0{Colors.RESET}")
    print(f"{Colors.CYAN}{'='*65}{Colors.RESET}")
    print(f"{Colors.YELLOW}  逐步测试定位问题 | 单项修复验证效果{Colors.RESET}\n")


def print_section(title: str):
    print(f"\n{Colors.BLUE}{'─'*65}{Colors.RESET}")
    print(f"  {Colors.BOLD}{Colors.CYAN}【{title}】{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*65}{Colors.RESET}")


# ============== 网络测试函数 ==============

def test_http_connectivity() -> Tuple[bool, float, str]:
    """测试 HTTP 连通性，返回 (是否成功, 响应时间ms, 详情)"""
    try:
        start = time.perf_counter()
        req = urllib.request.Request(
            "http://www.baidu.com",
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        with urllib.request.urlopen(req, timeout=10) as response:
            _ = response.read(1024)
        elapsed = (time.perf_counter() - start) * 1000
        return True, elapsed, f"响应时间: {elapsed:.0f}ms"
    except urllib.error.URLError as e:
        return False, -1, f"错误: {str(e.reason)[:50]}"
    except Exception as e:
        return False, -1, f"错误: {str(e)[:50]}"


def test_dns_resolution() -> Tuple[bool, float, str]:
    """测试 DNS 解析，返回 (是否成功, 解析时间ms, 详情)"""
    try:
        start = time.perf_counter()
        ip = socket.gethostbyname("www.baidu.com")
        elapsed = (time.perf_counter() - start) * 1000
        return True, elapsed, f"解析结果: {ip}, 耗时: {elapsed:.0f}ms"
    except socket.gaierror as e:
        return False, -1, f"DNS 解析失败: {e}"


def test_ping() -> Tuple[bool, float, str]:
    """测试 Ping，返回 (是否成功, 延迟ms, 详情)"""
    ret, stdout, _ = run_command("ping -n 1 -w 3000 114.114.114.114", timeout=5)
    if ret == 0 and ("TTL=" in stdout or "ttl=" in stdout.lower()):
        # 提取延迟
        for part in stdout.split():
            if "时间=" in part or "time=" in part.lower():
                try:
                    time_val = float(part.split('=')[1].replace('ms', '').replace('毫秒', ''))
                    return True, time_val, f"延迟: {time_val:.0f}ms"
                except:
                    pass
        return True, 0, "Ping 成功"
    return False, -1, "Ping 失败"


def run_network_test() -> Dict:
    """运行完整的网络测试"""
    results = {}

    print(f"\n  {Colors.CYAN}→{Colors.RESET} 测试 Ping...", end="", flush=True)
    ping_ok, ping_time, ping_detail = test_ping()
    results['ping'] = {'ok': ping_ok, 'time': ping_time, 'detail': ping_detail}
    print(f" {Colors.GREEN if ping_ok else Colors.RED}{ping_detail}{Colors.RESET}")

    print(f"  {Colors.CYAN}→{Colors.RESET} 测试 DNS 解析...", end="", flush=True)
    dns_ok, dns_time, dns_detail = test_dns_resolution()
    results['dns'] = {'ok': dns_ok, 'time': dns_time, 'detail': dns_detail}
    print(f" {Colors.GREEN if dns_ok else Colors.RED}{dns_detail}{Colors.RESET}")

    print(f"  {Colors.CYAN}→{Colors.RESET} 测试 HTTP 连通...", end="", flush=True)
    http_ok, http_time, http_detail = test_http_connectivity()
    results['http'] = {'ok': http_ok, 'time': http_time, 'detail': http_detail}
    print(f" {Colors.GREEN if http_ok else Colors.RED}{http_detail}{Colors.RESET}")

    return results


# ============== 问题检测函数 ==============

@dataclass
class ProblemInfo:
    """问题信息"""
    name: str
    detected: bool
    current_value: str
    description: str
    fix_action: str
    fix_function: callable


def check_ie_proxy() -> ProblemInfo:
    """检查 IE/系统代理"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            0, winreg.KEY_READ
        )
        proxy_enable = 0
        proxy_server = ""
        try:
            proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)

        detected = proxy_enable == 1 and proxy_server
        return ProblemInfo(
            name="系统代理 (IE/Edge)",
            detected=detected,
            current_value=f"{'启用' if proxy_enable else '禁用'}, 服务器: {proxy_server if proxy_server else '无'}",
            description="系统代理已启用，可能导致浏览器无法联网",
            fix_action="禁用系统代理",
            fix_function=fix_ie_proxy
        )
    except Exception as e:
        return ProblemInfo("系统代理", False, f"检测失败: {e}", "", "", None)


def check_pac_script() -> ProblemInfo:
    """检查 PAC 自动配置脚本"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            0, winreg.KEY_READ
        )
        auto_config_url = ""
        try:
            auto_config_url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)

        detected = bool(auto_config_url)
        return ProblemInfo(
            name="PAC 自动配置脚本",
            detected=detected,
            current_value=auto_config_url if auto_config_url else "未设置",
            description="PAC 脚本可能指向无效地址或已失效",
            fix_action="删除 PAC 脚本设置",
            fix_function=fix_pac_script
        )
    except Exception as e:
        return ProblemInfo("PAC 脚本", False, f"检测失败: {e}", "", "", None)


def check_winhttp_proxy() -> ProblemInfo:
    """检查 WinHTTP 代理"""
    ret, stdout, _ = run_command("netsh winhttp show proxy")
    has_proxy = "代理服务器" in stdout or ("Proxy Server" in stdout and "Direct access" not in stdout)

    return ProblemInfo(
        name="WinHTTP 代理",
        detected=has_proxy,
        current_value="已设置代理" if has_proxy else "直接访问",
        description="WinHTTP 代理会影响系统级 HTTP 请求",
        fix_action="重置 WinHTTP 代理",
        fix_function=fix_winhttp_proxy
    )


def check_winsock() -> ProblemInfo:
    """检查 Winsock 状态（简化检测）"""
    ret, stdout, _ = run_command("netsh winsock show catalog")

    suspicious_keywords = ['proxy', 'vpn', 'hook', 'inject', 'filter', 'tunnel']
    found = []
    for kw in suspicious_keywords:
        if kw.lower() in stdout.lower():
            found.append(kw)

    detected = len(found) > 0
    return ProblemInfo(
        name="Winsock 目录 (LSP)",
        detected=detected,
        current_value=f"发现可疑条目: {', '.join(found)}" if found else "正常",
        description="第三方软件可能污染了 Winsock 目录",
        fix_action="重置 Winsock 目录 (需重启)",
        fix_function=fix_winsock
    )


def check_dns_cache() -> ProblemInfo:
    """检查 DNS 缓存"""
    ret, stdout, _ = run_command("ipconfig /displaydns", timeout=10)
    count = stdout.count("记录名称") + stdout.count("Record Name")

    detected = count > 500  # 超过500条认为过多
    return ProblemInfo(
        name="DNS 缓存",
        detected=detected,
        current_value=f"{count} 条记录",
        description="DNS 缓存过多可能包含过期或错误记录",
        fix_action="清空 DNS 缓存",
        fix_function=fix_dns_cache
    )


# ============== 修复函数 ==============

def fix_ie_proxy() -> Tuple[bool, str]:
    """禁用 IE/系统代理"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "ProxyEnable", 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)
        return True, "系统代理已禁用"
    except Exception as e:
        return False, f"修复失败: {e}"


def fix_pac_script() -> Tuple[bool, str]:
    """删除 PAC 脚本"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            0, winreg.KEY_SET_VALUE
        )
        try:
            winreg.DeleteValue(key, "AutoConfigURL")
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
        return True, "PAC 脚本设置已删除"
    except Exception as e:
        return False, f"修复失败: {e}"


def fix_winhttp_proxy() -> Tuple[bool, str]:
    """重置 WinHTTP 代理"""
    ret, stdout, stderr = run_command("netsh winhttp reset proxy")
    if ret == 0:
        return True, "WinHTTP 代理已重置"
    return False, f"重置失败: {stderr}"


def fix_winsock() -> Tuple[bool, str]:
    """重置 Winsock"""
    ret, stdout, stderr = run_command("netsh winsock reset")
    if ret == 0:
        return True, "Winsock 已重置，需要重启计算机生效"
    return False, f"重置失败: {stderr}"


def fix_dns_cache() -> Tuple[bool, str]:
    """清空 DNS 缓存"""
    ret, stdout, stderr = run_command("ipconfig /flushdns")
    if ret == 0:
        return True, "DNS 缓存已清空"
    return False, f"清空失败: {stderr}"


# ============== 智能诊断流程 ==============

def smart_diagnose():
    """智能诊断：逐步测试定位问题"""
    print_section("第一步：测试当前网络状态")

    initial_test = run_network_test()

    # 判断问题类型
    ping_ok = initial_test['ping']['ok']
    dns_ok = initial_test['dns']['ok']
    http_ok = initial_test['http']['ok']

    if http_ok:
        print(f"\n  {Colors.GREEN}✓ 网络正常！HTTP 连通性测试通过{Colors.RESET}")
        print(f"  {Colors.YELLOW}提示：如果问题是间歇性的，请在问题发生时再次运行诊断{Colors.RESET}")
        return

    print(f"\n  {Colors.RED}✗ 检测到网络问题{Colors.RESET}")

    # 分析问题层次
    if not ping_ok:
        print(f"  {Colors.YELLOW}→ Ping 失败：可能是底层网络连接问题{Colors.RESET}")
    elif not dns_ok:
        print(f"  {Colors.YELLOW}→ DNS 解析失败：问题可能在 DNS 层{Colors.RESET}")
    else:
        print(f"  {Colors.YELLOW}→ HTTP 失败但 Ping/DNS 正常：问题可能在代理或应用层{Colors.RESET}")

    # 检测各项配置
    print_section("第二步：检测可能的问题原因")

    problems = [
        check_ie_proxy(),
        check_pac_script(),
        check_winhttp_proxy(),
        check_winsock(),
        check_dns_cache(),
    ]

    detected_problems = []
    for p in problems:
        status = f"{Colors.RED}[异常]{Colors.RESET}" if p.detected else f"{Colors.GREEN}[正常]{Colors.RESET}"
        print(f"\n  {status} {p.name}")
        print(f"       当前状态: {p.current_value}")
        if p.detected:
            print(f"       {Colors.YELLOW}可能原因: {p.description}{Colors.RESET}")
            detected_problems.append(p)

    if not detected_problems:
        print(f"\n  {Colors.YELLOW}未检测到明显的配置问题{Colors.RESET}")
        print(f"  可能原因：")
        print(f"    - 网络服务商问题")
        print(f"    - 防火墙或安全软件拦截")
        print(f"    - 路由器/网关问题")
        return

    # 逐一尝试修复
    print_section("第三步：逐一测试修复")
    print(f"\n  发现 {len(detected_problems)} 个可能的问题，将逐一尝试修复并测试效果")

    for i, problem in enumerate(detected_problems, 1):
        print(f"\n  {Colors.CYAN}━━━ 测试 {i}/{len(detected_problems)}: {problem.name} ━━━{Colors.RESET}")
        print(f"  操作: {problem.fix_action}")

        confirm = input(f"  是否执行此修复? (y=执行, n=跳过, q=退出): ").strip().lower()

        if confirm == 'q':
            print(f"\n  {Colors.YELLOW}诊断已中止{Colors.RESET}")
            return
        elif confirm != 'y':
            print(f"  已跳过")
            continue

        # 执行修复
        if problem.fix_function:
            success, message = problem.fix_function()
            if success:
                print(f"  {Colors.GREEN}✓ {message}{Colors.RESET}")
            else:
                print(f"  {Colors.RED}✗ {message}{Colors.RESET}")
                continue

        # 测试修复效果
        print(f"\n  {Colors.CYAN}测试修复效果...{Colors.RESET}")
        time.sleep(1)  # 等待系统生效

        test_result = run_network_test()

        if test_result['http']['ok']:
            print(f"\n  {Colors.GREEN}{'='*50}{Colors.RESET}")
            print(f"  {Colors.GREEN}✓ 问题已解决！{Colors.RESET}")
            print(f"  {Colors.GREEN}{'='*50}{Colors.RESET}")
            print(f"\n  {Colors.BOLD}诊断结论:{Colors.RESET}")
            print(f"  {Colors.WHITE}问题原因: {problem.name}{Colors.RESET}")
            print(f"  {Colors.WHITE}解决方法: {problem.fix_action}{Colors.RESET}")

            # 记录到日志
            save_diagnosis_result(problem)
            return
        else:
            print(f"  {Colors.YELLOW}此项修复未能解决问题，继续测试下一项...{Colors.RESET}")
            # 如果是代理设置，可能需要恢复（可选）

    print(f"\n  {Colors.YELLOW}所有已知项目测试完毕，问题仍未解决{Colors.RESET}")
    print(f"  建议：")
    print(f"    1. 检查防火墙或安全软件设置")
    print(f"    2. 尝试重启计算机（如果修改了 Winsock）")
    print(f"    3. 联系网络管理员或服务商")


def save_diagnosis_result(problem: ProblemInfo):
    """保存诊断结果到文件"""
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("diagnosis_log.txt", "a", encoding="utf-8") as f:
            f.write(f"\n[{timestamp}]\n")
            f.write(f"问题: {problem.name}\n")
            f.write(f"原状态: {problem.current_value}\n")
            f.write(f"修复操作: {problem.fix_action}\n")
            f.write("-" * 40 + "\n")
        print(f"\n  {Colors.DIM}诊断结果已保存到 diagnosis_log.txt{Colors.RESET}")
    except:
        pass


def single_item_test():
    """单项测试模式：选择单个项目进行修复和测试"""
    print_section("单项修复测试")

    print("\n  可测试的项目：")
    print(f"  {Colors.WHITE}1{Colors.RESET}. 禁用系统代理 (IE/Edge)")
    print(f"  {Colors.WHITE}2{Colors.RESET}. 删除 PAC 脚本")
    print(f"  {Colors.WHITE}3{Colors.RESET}. 重置 WinHTTP 代理")
    print(f"  {Colors.WHITE}4{Colors.RESET}. 清空 DNS 缓存")
    print(f"  {Colors.WHITE}5{Colors.RESET}. 重置 Winsock (需重启)")
    print(f"  {Colors.WHITE}0{Colors.RESET}. 返回")

    choice = input(f"\n  请选择 [0-5]: ").strip()

    fix_map = {
        '1': ("禁用系统代理", fix_ie_proxy),
        '2': ("删除 PAC 脚本", fix_pac_script),
        '3': ("重置 WinHTTP 代理", fix_winhttp_proxy),
        '4': ("清空 DNS 缓存", fix_dns_cache),
        '5': ("重置 Winsock", fix_winsock),
    }

    if choice not in fix_map:
        return

    name, fix_func = fix_map[choice]

    # 修复前测试
    print(f"\n  {Colors.CYAN}修复前网络状态:{Colors.RESET}")
    before = run_network_test()

    # 执行修复
    print(f"\n  {Colors.CYAN}执行: {name}...{Colors.RESET}")
    success, message = fix_func()
    print(f"  {'✓' if success else '✗'} {message}")

    if not success:
        return

    # 修复后测试
    time.sleep(1)
    print(f"\n  {Colors.CYAN}修复后网络状态:{Colors.RESET}")
    after = run_network_test()

    # 对比结果
    print(f"\n  {Colors.BOLD}效果对比:{Colors.RESET}")

    def compare(name, before_ok, after_ok, before_time, after_time):
        if before_ok and after_ok:
            diff = before_time - after_time
            if diff > 10:
                print(f"    {name}: {Colors.GREEN}提升 {diff:.0f}ms{Colors.RESET}")
            elif diff < -10:
                print(f"    {name}: {Colors.RED}下降 {-diff:.0f}ms{Colors.RESET}")
            else:
                print(f"    {name}: 无明显变化")
        elif not before_ok and after_ok:
            print(f"    {name}: {Colors.GREEN}已恢复正常{Colors.RESET}")
        elif before_ok and not after_ok:
            print(f"    {name}: {Colors.RED}变为异常{Colors.RESET}")
        else:
            print(f"    {name}: 仍然异常")

    compare("HTTP", before['http']['ok'], after['http']['ok'],
            before['http']['time'], after['http']['time'])
    compare("DNS", before['dns']['ok'], after['dns']['ok'],
            before['dns']['time'], after['dns']['time'])
    compare("Ping", before['ping']['ok'], after['ping']['ok'],
            before['ping']['time'], after['ping']['time'])

    if not before['http']['ok'] and after['http']['ok']:
        print(f"\n  {Colors.GREEN}{'='*50}{Colors.RESET}")
        print(f"  {Colors.GREEN}✓ 此操作解决了问题！{Colors.RESET}")
        print(f"  {Colors.GREEN}问题原因: {name} 相关配置异常{Colors.RESET}")
        print(f"  {Colors.GREEN}{'='*50}{Colors.RESET}")


def show_current_status():
    """显示当前网络配置状态"""
    print_section("当前网络配置状态")

    problems = [
        check_ie_proxy(),
        check_pac_script(),
        check_winhttp_proxy(),
        check_winsock(),
        check_dns_cache(),
    ]

    for p in problems:
        status = f"{Colors.RED}[异常]{Colors.RESET}" if p.detected else f"{Colors.GREEN}[正常]{Colors.RESET}"
        print(f"\n  {status} {p.name}")
        print(f"       {p.current_value}")


def show_menu():
    print(f"\n{Colors.CYAN}{'─'*65}{Colors.RESET}")
    print(f"{Colors.BOLD}请选择操作:{Colors.RESET}")
    print(f"  {Colors.WHITE}1{Colors.RESET}. 智能诊断 (自动逐步测试定位问题)")
    print(f"  {Colors.WHITE}2{Colors.RESET}. 单项测试 (选择单个修复项测试效果)")
    print(f"  {Colors.WHITE}3{Colors.RESET}. 仅测试网络 (不修复)")
    print(f"  {Colors.WHITE}4{Colors.RESET}. 查看当前配置状态")
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
        print(f"{Colors.YELLOW}请右键点击程序，选择「以管理员身份运行」{Colors.RESET}\n")

    while True:
        choice = show_menu()

        if choice == '0':
            print(f"\n{Colors.CYAN}感谢使用！{Colors.RESET}\n")
            break
        elif choice == '1':
            smart_diagnose()
        elif choice == '2':
            single_item_test()
        elif choice == '3':
            print_section("网络连通性测试")
            run_network_test()
        elif choice == '4':
            show_current_status()
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
