# Windows 网络工具箱

解决 Windows 系统常见网络问题的工具集。

## 包含工具

| 工具 | 路径 | 解决问题 |
|------|------|----------|
| **网络诊断修复** | 根目录 | 浏览器无网络但 QQ/Telegram 正常 |
| **网络监控优化** | `/monitor` | 长时间运行后网页变慢 |

---

# 工具一：网络诊断修复工具 (Network Doctor)

专门解决 **浏览器无网络但 QQ/Telegram 等通讯软件正常** 的问题。

## 问题症状

- 浏览器（Chrome、Edge、Firefox 等）无法访问网页
- QQ、Telegram、微信等通讯软件网络正常
- DNS 刷新、网络重置无效
- 重启后可能暂时恢复，但问题会复发

## 可能原因

| 原因 | 说明 |
|------|------|
| **系统代理被篡改** | 某些软件设置了代理但未正确清理 |
| **Winsock 目录损坏** | 第三方软件安装的 LSP 导致网络栈异常 |
| **WinHTTP 代理设置** | 系统级 HTTP 代理配置异常 |
| **DNS 缓存/设置问题** | DNS 解析层面的故障 |
| **Hosts 文件被修改** | 某些软件修改了 hosts 文件 |

> 通讯软件正常是因为它们使用自己的连接方式，不依赖系统代理设置。

## 快速开始

### 方法 1：一键快速修复（推荐）

双击运行 `quick_fix.bat`，自动执行以下修复：
- 禁用系统代理
- 重置 WinHTTP 代理
- 刷新 DNS 缓存
- 重置 Winsock 和 TCP/IP 栈

### 方法 2：使用 PowerShell 版本（无需 Python）

双击运行 `run_doctor_ps.bat`，启动图形化诊断工具。

### 方法 3：使用 Python 版本（完整功能）

1. 确保已安装 [Python 3.6+](https://www.python.org/downloads/)
2. 双击运行 `run_doctor.bat`

## 工具说明

| 文件 | 说明 |
|------|------|
| `quick_fix.bat` | 一键快速修复脚本，解决大多数问题 |
| `disable_proxy.bat` | 仅禁用代理设置 |
| `run_doctor_ps.bat` | 启动 PowerShell 版诊断工具 |
| `run_doctor.bat` | 启动 Python 版诊断工具 |
| `NetworkDoctor.ps1` | PowerShell 版完整诊断工具 |
| `network_doctor.py` | Python 版完整诊断工具 |

## 功能特性

### 诊断功能
- 检测系统代理设置 (IE/Edge 代理)
- 检测 WinHTTP 代理设置
- 检测 Winsock 目录状态（第三方 LSP）
- 检测 DNS 设置和解析能力
- 检测 Hosts 文件异常
- 检测网络适配器状态
- 测试网络连通性（Ping 测试）
- 测试 HTTP 连通性（关键测试）

### 修复功能
- 禁用/重置系统代理
- 重置 WinHTTP 代理
- 刷新 DNS 缓存
- 重新注册 DNS
- 重置 Winsock 目录
- 重置 TCP/IP 栈
- 释放/更新 IP 地址

## 使用截图

```
============================================================
      Windows 网络诊断修复工具 (Network Doctor) v1.0
============================================================
  专门解决: 浏览器无网络但QQ/Telegram等通讯软件正常的问题

──────────────────────────────────────────────────
请选择操作:
  1. 运行网络诊断
  2. 快速修复 (推荐首先尝试)
  3. 完整修复 (需要重启)
  4. 诊断 + 自动修复
  5. 仅禁用系统代理
  6. 仅重置 Winsock
  0. 退出
──────────────────────────────────────────────────
```

## 常见问题

### Q: 修复后需要重启吗？
A: 快速修复通常不需要重启。完整修复中的 Winsock 和 TCP/IP 重置需要重启才能完全生效。

### Q: 修复后问题还会复发吗？
A: 如果问题是由某个软件引起的，可能会复发。建议：
1. 检查最近安装的软件
2. 检查是否有 VPN 或代理软件
3. 检查杀毒软件的网络保护功能

### Q: 这个工具安全吗？
A: 工具只执行标准的 Windows 网络修复命令，不会修改系统关键文件。所有操作都是 Windows 官方支持的网络故障排除步骤。

### Q: 运行需要管理员权限吗？
A: 是的，网络修复操作需要管理员权限。脚本会自动请求提升权限。

## 手动修复命令

如果你偏好手动操作，可以在管理员命令提示符中执行：

```batch
:: 禁用系统代理
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" /v ProxyEnable /t REG_DWORD /d 0 /f

:: 重置 WinHTTP 代理
netsh winhttp reset proxy

:: 刷新 DNS
ipconfig /flushdns
ipconfig /registerdns

:: 重置 Winsock（需要重启）
netsh winsock reset

:: 重置 TCP/IP（需要重启）
netsh int ip reset
```

## 技术原理

本工具针对的是 **系统级网络配置问题**。当浏览器等使用系统网络设置的程序无法联网，而 QQ、Telegram 等使用独立连接的程序正常时，通常问题出在：

1. **代理层**：系统代理设置（IE 代理、WinHTTP 代理）被某些程序设置后未清理
2. **网络栈层**：Winsock LSP 被第三方程序污染
3. **DNS 层**：DNS 缓存或配置异常

工具通过重置这些层的配置来恢复网络正常。

---

# 工具二：网络监控与优化工具 (Network Monitor)

专门解决 **长时间运行后网页变慢，重启后改善** 的问题。

## 问题分析

| 原因 | 说明 |
|------|------|
| **DNS 缓存膨胀** | DNS 缓存条目过多，查询效率下降 |
| **ARP 缓存过大** | 网络地址解析表膨胀 |
| **NetBIOS 缓存堆积** | Windows 网络名称缓存过多 |
| **TCP 连接泄漏** | 未关闭的连接占用资源 |

## 快速使用

### 立即优化
```
monitor/optimize_now.bat
```

### 设置定时自动优化（推荐）
```
monitor/setup_scheduled_task.bat
```

### 完整监控工具
```
monitor/run_monitor_ps.bat
```

## 功能特性

- **实时监控**：DNS解析时间、Ping延迟、HTTP响应时间
- **资源监控**：TCP连接数、缓存条目数、内存使用率
- **一键优化**：清理 DNS/ARP/NetBIOS 缓存
- **定时任务**：支持设置自动定期清理
- **性能对比**：优化前后效果对比

详细说明请查看 [monitor/README.md](monitor/README.md)

---

## 系统要求

- Windows 10 / Windows 11
- 管理员权限
- Python 3.6+（仅 Python 版本需要）

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
