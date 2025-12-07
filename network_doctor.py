#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 网络诊断修复工具 (Network Doctor)
用于检测和修复 Windows 系统中浏览器无网络但通讯软件正常的问题

作者: Network Doctor Team
版本: 1.0.0
"""

import subprocess
import sys
import os
import ctypes
import winreg
import socket
import json
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum


class DiagnosticStatus(Enum):
    """诊断状态枚举"""
    OK = "✓ 正常"
    WARNING = "⚠ 警告"
    ERROR = "✗ 异常"
    INFO = "ℹ 信息"


@dataclass
class DiagnosticResult:
    """诊断结果数据类"""
    name: str
    status: DiagnosticStatus
    message: str
    details: Optional[str] = None
    fix_available: bool = False
    fix_command: Optional[str] = None


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


def is_admin() -> bool:
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_command(command: str, timeout: int = 30) -> Tuple[int, str, str]:
    """
    执行系统命令并返回结果

    Args:
        command: 要执行的命令
        timeout: 超时时间(秒)

    Returns:
        (返回码, 标准输出, 标准错误)
    """
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


def print_header():
    """打印程序头部"""
    print(f"\n{Colors.CYAN}{'='*60}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.WHITE}      Windows 网络诊断修复工具 (Network Doctor) v1.0{Colors.RESET}")
    print(f"{Colors.CYAN}{'='*60}{Colors.RESET}")
    print(f"{Colors.YELLOW}专门解决: 浏览器无网络但QQ/Telegram等通讯软件正常的问题{Colors.RESET}\n")


def print_section(title: str):
    """打印分节标题"""
    print(f"\n{Colors.BLUE}{'─'*50}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.CYAN}【{title}】{Colors.RESET}")
    print(f"{Colors.BLUE}{'─'*50}{Colors.RESET}")


def print_result(result: DiagnosticResult):
    """打印诊断结果"""
    color = {
        DiagnosticStatus.OK: Colors.GREEN,
        DiagnosticStatus.WARNING: Colors.YELLOW,
        DiagnosticStatus.ERROR: Colors.RED,
        DiagnosticStatus.INFO: Colors.BLUE
    }.get(result.status, Colors.WHITE)

    print(f"\n  {color}{result.status.value}{Colors.RESET} {result.name}")
    print(f"      {result.message}")
    if result.details:
        for line in result.details.split('\n'):
            if line.strip():
                print(f"      {Colors.WHITE}{line}{Colors.RESET}")


class NetworkDiagnostics:
    """网络诊断类"""

    def __init__(self):
        self.results: List[DiagnosticResult] = []
        self.issues_found = 0

    def check_proxy_settings(self) -> DiagnosticResult:
        """检查系统代理设置 (IE/Edge 代理)"""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
                0, winreg.KEY_READ
            )

            proxy_enable = 0
            proxy_server = ""
            proxy_override = ""
            auto_config_url = ""

            try:
                proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            except FileNotFoundError:
                pass

            try:
                proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
            except FileNotFoundError:
                pass

            try:
                proxy_override, _ = winreg.QueryValueEx(key, "ProxyOverride")
            except FileNotFoundError:
                pass

            try:
                auto_config_url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
            except FileNotFoundError:
                pass

            winreg.CloseKey(key)

            details = f"代理启用: {'是' if proxy_enable else '否'}\n"
            details += f"代理服务器: {proxy_server if proxy_server else '无'}\n"
            details += f"PAC 脚本: {auto_config_url if auto_config_url else '无'}\n"
            details += f"代理例外: {proxy_override if proxy_override else '无'}"

            if proxy_enable and proxy_server:
                self.issues_found += 1
                return DiagnosticResult(
                    name="系统代理设置 (IE/Edge)",
                    status=DiagnosticStatus.WARNING,
                    message="检测到系统代理已启用，这可能导致部分程序无法联网",
                    details=details,
                    fix_available=True,
                    fix_command="disable_ie_proxy"
                )
            elif auto_config_url:
                return DiagnosticResult(
                    name="系统代理设置 (IE/Edge)",
                    status=DiagnosticStatus.INFO,
                    message="检测到 PAC 自动配置脚本，请确认是否需要",
                    details=details,
                    fix_available=True,
                    fix_command="disable_ie_proxy"
                )
            else:
                return DiagnosticResult(
                    name="系统代理设置 (IE/Edge)",
                    status=DiagnosticStatus.OK,
                    message="系统代理未启用",
                    details=details
                )

        except Exception as e:
            return DiagnosticResult(
                name="系统代理设置 (IE/Edge)",
                status=DiagnosticStatus.ERROR,
                message=f"无法读取代理设置: {str(e)}"
            )

    def check_winhttp_proxy(self) -> DiagnosticResult:
        """检查 WinHTTP 代理设置"""
        ret, stdout, stderr = run_command("netsh winhttp show proxy")

        details = stdout.strip() if stdout else stderr

        if "直接访问" in stdout or "Direct access" in stdout:
            return DiagnosticResult(
                name="WinHTTP 代理设置",
                status=DiagnosticStatus.OK,
                message="WinHTTP 设置为直接访问（无代理）",
                details=details
            )
        elif "代理服务器" in stdout or "Proxy Server" in stdout:
            self.issues_found += 1
            return DiagnosticResult(
                name="WinHTTP 代理设置",
                status=DiagnosticStatus.WARNING,
                message="检测到 WinHTTP 代理设置，可能导致系统级网络请求异常",
                details=details,
                fix_available=True,
                fix_command="reset_winhttp_proxy"
            )
        else:
            return DiagnosticResult(
                name="WinHTTP 代理设置",
                status=DiagnosticStatus.INFO,
                message="WinHTTP 代理状态",
                details=details
            )

    def check_winsock(self) -> DiagnosticResult:
        """检查 Winsock 目录状态"""
        ret, stdout, stderr = run_command("netsh winsock show catalog")

        if ret == 0 and stdout:
            # 统计 LSP 条目数量
            lsp_count = stdout.count("Winsock 目录") or stdout.count("Winsock Catalog")

            # 检查是否有可疑的第三方 LSP
            suspicious_keywords = ['proxy', 'vpn', 'hook', 'inject', 'filter']
            suspicious_found = []
            for keyword in suspicious_keywords:
                if keyword.lower() in stdout.lower():
                    suspicious_found.append(keyword)

            if suspicious_found:
                self.issues_found += 1
                return DiagnosticResult(
                    name="Winsock 目录",
                    status=DiagnosticStatus.WARNING,
                    message=f"检测到可能影响网络的第三方 LSP: {', '.join(suspicious_found)}",
                    details=f"共有 {lsp_count} 个 Winsock 目录条目",
                    fix_available=True,
                    fix_command="reset_winsock"
                )
            else:
                return DiagnosticResult(
                    name="Winsock 目录",
                    status=DiagnosticStatus.OK,
                    message="Winsock 目录看起来正常",
                    details=f"共有 {lsp_count} 个条目"
                )
        else:
            return DiagnosticResult(
                name="Winsock 目录",
                status=DiagnosticStatus.ERROR,
                message="无法获取 Winsock 目录信息",
                details=stderr,
                fix_available=True,
                fix_command="reset_winsock"
            )

    def check_dns_settings(self) -> DiagnosticResult:
        """检查 DNS 设置"""
        ret, stdout, stderr = run_command("ipconfig /all")

        dns_servers = []
        lines = stdout.split('\n')
        for i, line in enumerate(lines):
            if "DNS 服务器" in line or "DNS Servers" in line:
                # 提取 DNS 服务器地址
                parts = line.split(':')
                if len(parts) > 1:
                    dns = parts[1].strip()
                    if dns:
                        dns_servers.append(dns)
                # 检查下一行是否还有 DNS 地址
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and not any(x in next_line for x in [':', '适配器', 'adapter']):
                        dns_servers.append(next_line)

        details = f"检测到的 DNS 服务器: {', '.join(dns_servers) if dns_servers else '无'}"

        if not dns_servers:
            self.issues_found += 1
            return DiagnosticResult(
                name="DNS 设置",
                status=DiagnosticStatus.ERROR,
                message="未检测到 DNS 服务器配置",
                details=details,
                fix_available=True,
                fix_command="reset_dns"
            )
        else:
            return DiagnosticResult(
                name="DNS 设置",
                status=DiagnosticStatus.OK,
                message=f"已配置 {len(dns_servers)} 个 DNS 服务器",
                details=details
            )

    def check_dns_resolution(self) -> DiagnosticResult:
        """检查 DNS 解析功能"""
        test_domains = [
            ("www.baidu.com", "百度"),
            ("www.qq.com", "腾讯"),
            ("www.google.com", "Google"),
        ]

        results = []
        failed = 0

        for domain, name in test_domains:
            try:
                ip = socket.gethostbyname(domain)
                results.append(f"{name}({domain}): {ip}")
            except socket.gaierror:
                results.append(f"{name}({domain}): 解析失败")
                failed += 1

        details = '\n'.join(results)

        if failed == len(test_domains):
            self.issues_found += 1
            return DiagnosticResult(
                name="DNS 解析测试",
                status=DiagnosticStatus.ERROR,
                message="DNS 解析完全失败，这可能是主要问题",
                details=details,
                fix_available=True,
                fix_command="flush_dns"
            )
        elif failed > 0:
            return DiagnosticResult(
                name="DNS 解析测试",
                status=DiagnosticStatus.WARNING,
                message=f"部分域名解析失败 ({failed}/{len(test_domains)})",
                details=details,
                fix_available=True,
                fix_command="flush_dns"
            )
        else:
            return DiagnosticResult(
                name="DNS 解析测试",
                status=DiagnosticStatus.OK,
                message="DNS 解析正常",
                details=details
            )

    def check_hosts_file(self) -> DiagnosticResult:
        """检查 hosts 文件"""
        hosts_path = r"C:\Windows\System32\drivers\etc\hosts"

        try:
            with open(hosts_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            lines = [l.strip() for l in content.split('\n') if l.strip() and not l.strip().startswith('#')]

            # 检查是否有可疑的条目
            suspicious = []
            common_sites = ['google', 'facebook', 'youtube', 'twitter', 'github']

            for line in lines:
                for site in common_sites:
                    if site in line.lower():
                        suspicious.append(line)
                        break

            details = f"自定义条目数: {len(lines)}\n"
            if suspicious:
                details += f"可疑条目:\n" + '\n'.join(suspicious[:5])

            if suspicious:
                self.issues_found += 1
                return DiagnosticResult(
                    name="Hosts 文件",
                    status=DiagnosticStatus.WARNING,
                    message=f"检测到 {len(suspicious)} 个可能影响网络的 hosts 条目",
                    details=details,
                    fix_available=True,
                    fix_command="backup_hosts"
                )
            elif len(lines) > 50:
                return DiagnosticResult(
                    name="Hosts 文件",
                    status=DiagnosticStatus.WARNING,
                    message=f"Hosts 文件条目较多 ({len(lines)} 条)",
                    details=details
                )
            else:
                return DiagnosticResult(
                    name="Hosts 文件",
                    status=DiagnosticStatus.OK,
                    message="Hosts 文件正常",
                    details=details
                )

        except Exception as e:
            return DiagnosticResult(
                name="Hosts 文件",
                status=DiagnosticStatus.ERROR,
                message=f"无法读取 hosts 文件: {str(e)}"
            )

    def check_network_adapters(self) -> DiagnosticResult:
        """检查网络适配器状态"""
        ret, stdout, stderr = run_command("ipconfig")

        # 检查是否有活动的网络适配器
        adapters = []
        current_adapter = None
        has_ip = False

        for line in stdout.split('\n'):
            if '适配器' in line or 'adapter' in line.lower():
                if current_adapter and has_ip:
                    adapters.append(current_adapter)
                current_adapter = line.strip().rstrip(':')
                has_ip = False
            elif 'IPv4' in line or 'IP 地址' in line:
                has_ip = True

        if current_adapter and has_ip:
            adapters.append(current_adapter)

        details = f"活动的网络适配器: {len(adapters)}\n"
        details += '\n'.join(adapters[:5])

        if not adapters:
            self.issues_found += 1
            return DiagnosticResult(
                name="网络适配器",
                status=DiagnosticStatus.ERROR,
                message="未检测到活动的网络适配器",
                details=details,
                fix_available=True,
                fix_command="reset_adapter"
            )
        else:
            return DiagnosticResult(
                name="网络适配器",
                status=DiagnosticStatus.OK,
                message=f"检测到 {len(adapters)} 个活动的网络适配器",
                details=details
            )

    def check_connectivity(self) -> DiagnosticResult:
        """检查网络连通性"""
        test_targets = [
            ("114.114.114.114", "国内 DNS"),
            ("223.5.5.5", "阿里 DNS"),
            ("8.8.8.8", "Google DNS"),
        ]

        results = []
        success = 0

        for ip, name in test_targets:
            ret, stdout, stderr = run_command(f"ping -n 1 -w 2000 {ip}")
            if ret == 0 and ("TTL=" in stdout or "ttl=" in stdout.lower()):
                results.append(f"{name}({ip}): 可达")
                success += 1
            else:
                results.append(f"{name}({ip}): 不可达")

        details = '\n'.join(results)

        if success == 0:
            self.issues_found += 1
            return DiagnosticResult(
                name="网络连通性",
                status=DiagnosticStatus.ERROR,
                message="无法连接到任何测试目标，请检查网络连接",
                details=details
            )
        elif success < len(test_targets):
            return DiagnosticResult(
                name="网络连通性",
                status=DiagnosticStatus.WARNING,
                message=f"部分目标可达 ({success}/{len(test_targets)})",
                details=details
            )
        else:
            return DiagnosticResult(
                name="网络连通性",
                status=DiagnosticStatus.OK,
                message="网络连通性正常",
                details=details
            )

    def check_http_connectivity(self) -> DiagnosticResult:
        """检查 HTTP 连通性（这是关键测试）"""
        import urllib.request
        import ssl

        test_urls = [
            ("http://www.baidu.com", "百度 HTTP"),
            ("https://www.baidu.com", "百度 HTTPS"),
            ("http://www.qq.com", "腾讯 HTTP"),
        ]

        results = []
        success = 0

        # 创建不验证 SSL 的上下文（仅用于测试）
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        for url, name in test_urls:
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10, context=ctx if url.startswith('https') else None) as response:
                    if response.status == 200:
                        results.append(f"{name}: 正常 (状态码: {response.status})")
                        success += 1
                    else:
                        results.append(f"{name}: 异常 (状态码: {response.status})")
            except urllib.error.URLError as e:
                results.append(f"{name}: 失败 ({str(e.reason)[:30]})")
            except Exception as e:
                results.append(f"{name}: 失败 ({str(e)[:30]})")

        details = '\n'.join(results)

        if success == 0:
            self.issues_found += 1
            return DiagnosticResult(
                name="HTTP 连通性测试",
                status=DiagnosticStatus.ERROR,
                message="HTTP 请求全部失败！这是问题的关键表现",
                details=details,
                fix_available=True,
                fix_command="full_reset"
            )
        elif success < len(test_urls):
            self.issues_found += 1
            return DiagnosticResult(
                name="HTTP 连通性测试",
                status=DiagnosticStatus.WARNING,
                message=f"部分 HTTP 请求失败 ({success}/{len(test_urls)})",
                details=details,
                fix_available=True,
                fix_command="full_reset"
            )
        else:
            return DiagnosticResult(
                name="HTTP 连通性测试",
                status=DiagnosticStatus.OK,
                message="HTTP 连通性正常",
                details=details
            )

    def check_firewall_status(self) -> DiagnosticResult:
        """检查 Windows 防火墙状态"""
        ret, stdout, stderr = run_command("netsh advfirewall show allprofiles state")

        details = stdout.strip() if stdout else stderr

        if ret == 0:
            return DiagnosticResult(
                name="Windows 防火墙",
                status=DiagnosticStatus.INFO,
                message="防火墙状态信息",
                details=details
            )
        else:
            return DiagnosticResult(
                name="Windows 防火墙",
                status=DiagnosticStatus.WARNING,
                message="无法获取防火墙状态",
                details=stderr
            )

    def run_all_diagnostics(self) -> List[DiagnosticResult]:
        """运行所有诊断"""
        print_section("开始网络诊断")

        diagnostics = [
            ("检查系统代理设置...", self.check_proxy_settings),
            ("检查 WinHTTP 代理...", self.check_winhttp_proxy),
            ("检查 Winsock 目录...", self.check_winsock),
            ("检查 DNS 设置...", self.check_dns_settings),
            ("测试 DNS 解析...", self.check_dns_resolution),
            ("检查 Hosts 文件...", self.check_hosts_file),
            ("检查网络适配器...", self.check_network_adapters),
            ("测试网络连通性...", self.check_connectivity),
            ("测试 HTTP 连通性...", self.check_http_connectivity),
            ("检查防火墙状态...", self.check_firewall_status),
        ]

        for msg, func in diagnostics:
            print(f"\n  {Colors.CYAN}→{Colors.RESET} {msg}", end="", flush=True)
            result = func()
            self.results.append(result)
            print(f" {Colors.GREEN}完成{Colors.RESET}")

        return self.results


class NetworkRepair:
    """网络修复类"""

    @staticmethod
    def disable_ie_proxy() -> bool:
        """禁用 IE/系统代理"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在禁用系统代理...")
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
                0, winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(key, "ProxyEnable", 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, "ProxyServer", 0, winreg.REG_SZ, "")
            # 删除 PAC 脚本设置
            try:
                winreg.DeleteValue(key, "AutoConfigURL")
            except FileNotFoundError:
                pass
            winreg.CloseKey(key)
            print(f"  {Colors.GREEN}✓{Colors.RESET} 系统代理已禁用")
            return True
        except Exception as e:
            print(f"  {Colors.RED}✗{Colors.RESET} 禁用系统代理失败: {e}")
            return False

    @staticmethod
    def reset_winhttp_proxy() -> bool:
        """重置 WinHTTP 代理"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在重置 WinHTTP 代理...")
        ret, stdout, stderr = run_command("netsh winhttp reset proxy")
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} WinHTTP 代理已重置")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.RESET} 重置失败: {stderr}")
            return False

    @staticmethod
    def reset_winsock() -> bool:
        """重置 Winsock 目录"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在重置 Winsock 目录...")
        ret, stdout, stderr = run_command("netsh winsock reset")
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} Winsock 已重置 (需要重启生效)")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.RESET} 重置失败: {stderr}")
            return False

    @staticmethod
    def flush_dns() -> bool:
        """刷新 DNS 缓存"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在刷新 DNS 缓存...")
        ret, stdout, stderr = run_command("ipconfig /flushdns")
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} DNS 缓存已刷新")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.RESET} 刷新失败: {stderr}")
            return False

    @staticmethod
    def reset_dns() -> bool:
        """重置 DNS 设置"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在注册 DNS...")
        ret, stdout, stderr = run_command("ipconfig /registerdns")
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} DNS 已重新注册")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.RESET} 注册失败: {stderr}")
            return False

    @staticmethod
    def reset_tcp_ip() -> bool:
        """重置 TCP/IP 栈"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在重置 TCP/IP 栈...")
        ret, stdout, stderr = run_command("netsh int ip reset")
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} TCP/IP 栈已重置 (需要重启生效)")
            return True
        else:
            print(f"  {Colors.RED}✗{Colors.RESET} 重置失败: {stderr}")
            return False

    @staticmethod
    def reset_adapter() -> bool:
        """重置网络适配器"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在重置网络适配器...")
        # 禁用再启用所有网络适配器
        ret1, _, _ = run_command("netsh interface set interface \"以太网\" disable", timeout=10)
        ret2, _, _ = run_command("netsh interface set interface \"以太网\" enable", timeout=10)
        ret3, _, _ = run_command("netsh interface set interface \"WLAN\" disable", timeout=10)
        ret4, _, _ = run_command("netsh interface set interface \"WLAN\" enable", timeout=10)
        print(f"  {Colors.GREEN}✓{Colors.RESET} 网络适配器重置命令已执行")
        return True

    @staticmethod
    def release_renew_ip() -> bool:
        """释放和更新 IP 地址"""
        print(f"\n  {Colors.CYAN}→{Colors.RESET} 正在释放 IP 地址...")
        run_command("ipconfig /release", timeout=30)
        print(f"  {Colors.CYAN}→{Colors.RESET} 正在更新 IP 地址...")
        ret, stdout, stderr = run_command("ipconfig /renew", timeout=60)
        if ret == 0:
            print(f"  {Colors.GREEN}✓{Colors.RESET} IP 地址已更新")
            return True
        else:
            print(f"  {Colors.YELLOW}⚠{Colors.RESET} IP 更新可能未完全成功")
            return False

    def full_repair(self) -> bool:
        """执行完整修复"""
        print_section("开始网络修复")

        success_count = 0
        total = 7

        # 1. 禁用系统代理
        if self.disable_ie_proxy():
            success_count += 1

        # 2. 重置 WinHTTP 代理
        if self.reset_winhttp_proxy():
            success_count += 1

        # 3. 刷新 DNS
        if self.flush_dns():
            success_count += 1

        # 4. 重置 DNS
        if self.reset_dns():
            success_count += 1

        # 5. 重置 Winsock
        if self.reset_winsock():
            success_count += 1

        # 6. 重置 TCP/IP
        if self.reset_tcp_ip():
            success_count += 1

        # 7. 释放和更新 IP
        if self.release_renew_ip():
            success_count += 1

        return success_count >= total // 2

    def quick_repair(self) -> bool:
        """快速修复（只修复最可能的问题）"""
        print_section("开始快速修复")

        success_count = 0

        # 1. 禁用系统代理（最常见原因）
        if self.disable_ie_proxy():
            success_count += 1

        # 2. 重置 WinHTTP 代理
        if self.reset_winhttp_proxy():
            success_count += 1

        # 3. 刷新 DNS
        if self.flush_dns():
            success_count += 1

        return success_count >= 2


def show_summary(results: List[DiagnosticResult], issues_found: int):
    """显示诊断摘要"""
    print_section("诊断结果摘要")

    for result in results:
        print_result(result)

    print(f"\n{Colors.BLUE}{'─'*50}{Colors.RESET}")
    if issues_found > 0:
        print(f"\n  {Colors.YELLOW}发现 {issues_found} 个潜在问题{Colors.RESET}")
        print(f"  建议执行修复操作来解决网络问题")
    else:
        print(f"\n  {Colors.GREEN}未发现明显问题{Colors.RESET}")
        print(f"  如果问题持续，建议尝试完整修复")


def show_menu() -> str:
    """显示主菜单"""
    print(f"\n{Colors.CYAN}{'─'*50}{Colors.RESET}")
    print(f"{Colors.BOLD}请选择操作:{Colors.RESET}")
    print(f"  {Colors.WHITE}1{Colors.RESET}. 运行网络诊断")
    print(f"  {Colors.WHITE}2{Colors.RESET}. 快速修复 (推荐首先尝试)")
    print(f"  {Colors.WHITE}3{Colors.RESET}. 完整修复 (需要重启)")
    print(f"  {Colors.WHITE}4{Colors.RESET}. 诊断 + 自动修复")
    print(f"  {Colors.WHITE}5{Colors.RESET}. 仅禁用系统代理")
    print(f"  {Colors.WHITE}6{Colors.RESET}. 仅重置 Winsock")
    print(f"  {Colors.WHITE}7{Colors.RESET}. 导出诊断报告")
    print(f"  {Colors.WHITE}0{Colors.RESET}. 退出")
    print(f"{Colors.CYAN}{'─'*50}{Colors.RESET}")

    return input(f"\n{Colors.BOLD}请输入选项 [0-7]: {Colors.RESET}").strip()


def export_report(results: List[DiagnosticResult]):
    """导出诊断报告"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"network_diagnostic_report_{timestamp}.txt"

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("Windows 网络诊断报告\n")
        f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")

        for result in results:
            f.write(f"[{result.status.value}] {result.name}\n")
            f.write(f"  {result.message}\n")
            if result.details:
                f.write(f"  详情:\n")
                for line in result.details.split('\n'):
                    f.write(f"    {line}\n")
            f.write("\n")

    print(f"\n  {Colors.GREEN}✓{Colors.RESET} 报告已保存至: {filename}")


def main():
    """主函数"""
    # 启用 Windows 终端颜色支持
    if sys.platform == 'win32':
        os.system('color')
        # 尝试启用 ANSI 支持
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
        except:
            pass

    print_header()

    # 检查管理员权限
    if not is_admin():
        print(f"{Colors.RED}{'!'*50}{Colors.RESET}")
        print(f"{Colors.RED}警告: 程序未以管理员权限运行!{Colors.RESET}")
        print(f"{Colors.RED}部分修复功能可能无法正常工作{Colors.RESET}")
        print(f"{Colors.RED}请右键点击程序，选择「以管理员身份运行」{Colors.RESET}")
        print(f"{Colors.RED}{'!'*50}{Colors.RESET}")

    diagnostics = NetworkDiagnostics()
    repair = NetworkRepair()
    results = []

    while True:
        choice = show_menu()

        if choice == '0':
            print(f"\n{Colors.CYAN}感谢使用网络诊断修复工具！{Colors.RESET}\n")
            break

        elif choice == '1':
            results = diagnostics.run_all_diagnostics()
            show_summary(results, diagnostics.issues_found)

        elif choice == '2':
            repair.quick_repair()
            print(f"\n{Colors.GREEN}快速修复完成！{Colors.RESET}")
            print(f"{Colors.YELLOW}如果问题仍然存在，请尝试完整修复（选项3）{Colors.RESET}")

        elif choice == '3':
            repair.full_repair()
            print(f"\n{Colors.GREEN}完整修复完成！{Colors.RESET}")
            print(f"{Colors.YELLOW}某些更改需要重启计算机才能生效{Colors.RESET}")
            restart = input(f"\n是否立即重启计算机? (y/n): ").strip().lower()
            if restart == 'y':
                print("正在重启计算机...")
                run_command("shutdown /r /t 5")

        elif choice == '4':
            results = diagnostics.run_all_diagnostics()
            show_summary(results, diagnostics.issues_found)

            if diagnostics.issues_found > 0:
                print(f"\n{Colors.YELLOW}检测到问题，是否进行自动修复?{Colors.RESET}")
                confirm = input("输入 y 确认修复，其他键跳过: ").strip().lower()
                if confirm == 'y':
                    repair.full_repair()
                    print(f"\n{Colors.GREEN}修复完成！建议重启计算机{Colors.RESET}")

        elif choice == '5':
            repair.disable_ie_proxy()
            repair.reset_winhttp_proxy()
            print(f"\n{Colors.GREEN}代理设置已清除！{Colors.RESET}")

        elif choice == '6':
            repair.reset_winsock()
            print(f"\n{Colors.GREEN}Winsock 已重置，需要重启生效{Colors.RESET}")

        elif choice == '7':
            if not results:
                print(f"\n{Colors.YELLOW}请先运行诊断（选项1）{Colors.RESET}")
            else:
                export_report(results)

        else:
            print(f"\n{Colors.RED}无效选项，请重新选择{Colors.RESET}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}程序被用户中断{Colors.RESET}")
    except Exception as e:
        print(f"\n{Colors.RED}程序出错: {e}{Colors.RESET}")
        input("按回车键退出...")
