# 网络监控与优化工具 (Network Monitor)

专门解决 **长时间运行后网页变慢** 的问题。

## 问题分析

电脑长时间运行后网页变慢，重启后改善，通常是由以下原因导致：

| 原因 | 说明 |
|------|------|
| **DNS 缓存膨胀** | DNS 缓存条目过多，导致查询效率下降 |
| **ARP 缓存过大** | 网络地址解析表膨胀，增加解析延迟 |
| **NetBIOS 缓存堆积** | Windows 网络名称缓存过多 |
| **TCP 连接泄漏** | 未正确关闭的连接占用系统资源 |
| **内存资源耗尽** | 浏览器或系统进程内存泄漏 |

## 快速使用

### 方法 1：立即优化（推荐）

双击运行 `optimize_now.bat`，一键清理所有网络缓存。

### 方法 2：设置定时自动优化

双击运行 `setup_scheduled_task.bat`，设置自动定期清理。

建议间隔：
- 长时间运行的电脑：每 30 分钟
- 普通使用：每 1-2 小时

### 方法 3：使用完整监控工具

PowerShell 版本（无需 Python）：
```
run_monitor_ps.bat
```

Python 版本：
```
run_monitor.bat
```

## 工具文件说明

| 文件 | 说明 |
|------|------|
| `optimize_now.bat` | 立即执行网络优化 |
| `setup_scheduled_task.bat` | 设置定时自动优化任务 |
| `run_monitor.bat` | 启动 Python 版监控工具 |
| `run_monitor_ps.bat` | 启动 PowerShell 版监控工具 |
| `network_monitor.py` | Python 完整监控程序 |
| `NetworkMonitor.ps1` | PowerShell 完整监控程序 |

## 功能特性

### 监控功能
- **DNS 解析时间** - 监测域名解析速度
- **Ping 延迟** - 监测网络往返时间
- **HTTP 响应时间** - 监测网页加载速度（关键指标）
- **TCP 连接数** - 监测系统网络连接
- **DNS/ARP 缓存条目** - 监测缓存膨胀情况
- **内存使用率** - 监测系统资源

### 优化功能
- **清理 DNS 缓存** - 释放陈旧的 DNS 记录
- **清理 ARP 缓存** - 刷新网络地址映射
- **清理 NetBIOS 缓存** - 清理 Windows 网络名称缓存
- **优化 TCP 设置** - 调整网络传输参数
- **重新注册 DNS** - 刷新 DNS 客户端配置

### 性能对比
工具支持"优化前后对比测试"，直观展示优化效果：
- 记录优化前的网络指标
- 执行优化操作
- 记录优化后的指标
- 显示改善百分比

## 定时任务

通过 `setup_scheduled_task.bat` 可以创建 Windows 计划任务：

- 任务名称：`NetworkOptimizer`
- 执行间隔：可选 30分钟 / 1小时 / 2小时 / 4小时
- 执行方式：后台静默运行，不影响正常使用

删除任务：再次运行脚本，选择删除选项

## 手动优化命令

如需手动执行，在管理员命令提示符中运行：

```batch
:: 清理 DNS 缓存
ipconfig /flushdns

:: 清理 ARP 缓存
netsh interface ip delete arpcache
arp -d *

:: 清理 NetBIOS 缓存
nbtstat -R
nbtstat -RR

:: 重新注册 DNS
ipconfig /registerdns
```

## 使用建议

1. **首次使用**：运行 `optimize_now.bat` 立即清理
2. **长期方案**：运行 `setup_scheduled_task.bat` 设置自动清理
3. **问题排查**：使用 `run_monitor_ps.bat` 查看详细状态

## 系统要求

- Windows 10 / Windows 11
- 管理员权限
- Python 3.6+（仅 Python 版本需要）
