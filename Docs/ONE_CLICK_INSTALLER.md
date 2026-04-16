# GOWordAgent Linux 单文件安装器

## 概述

提供两种安装方式，满足不同用户需求：

| 方式 | 适用场景 | 文件大小 | 技术要求 |
|------|----------|----------|----------|
| **单文件安装器 (.run)** | 最终用户，双击安装 | ~80MB | 无 |
| **脚本安装 (tar.gz)** | 开发者，可定制 | ~75MB | 基本命令行 |

---

## 方式一：单文件安装器（推荐）

### 用户安装步骤

```bash
# 1. 下载安装器
wget https://your-server/GOWordAgent-Install-linux-x64-1.0.0.run

# 2. 添加执行权限
chmod +x GOWordAgent-Install-linux-x64-1.0.0.run

# 3. 运行安装（支持图形界面）
./GOWordAgent-Install-linux-x64-1.0.0.run
```

### 特点

- **双击安装**：支持桌面环境的双击执行
- **图形界面**：自动检测并使用 zenity 显示图形对话框
- **命令行界面**：无图形环境时自动回退到命令行
- **自动配置**：自动检测并使用 Systemd 或创建手动启动脚本
- **验证安装**：安装后自动验证服务状态

### 安装流程

```
[1/5] 检查系统架构    → 检查 x86_64
[2/5] 检查系统依赖    → 检查 tar
[3/5] 配置安装路径    → ~/.local/opt/gowordagent
[4/5] 解压安装文件    → 解压并复制文件
[5/5] 配置系统服务    → 启动后端服务
```

---

## 方式二：脚本安装

### 用户安装步骤

```bash
# 1. 下载并解压
tar -xzf gowordagent-linux-x64-1.0.0.tar.gz
cd gowordagent-linux-x64-1.0.0

# 2. 运行安装脚本
./deploy-linux.sh
```

---

## 构建安装器

### 前提条件

- Windows: 需要 WSL (Windows Subsystem for Linux)
- Linux: 任何 Linux 发行版

### 构建步骤

#### 在 Linux 上构建

```bash
# 1. 构建发布版本
./build-release.sh

# 2. 构建单文件安装器
./create-oneclick-installer.sh

# 输出: GOWordAgent-Install-linux-x64-1.0.0.run
```

#### 在 Windows 上构建（需要 WSL）

```powershell
# 1. 在 WSL 中构建发布版本
wsl ./build-release.sh

# 2. 使用 PowerShell 构建安装器
powershell -ExecutionPolicy Bypass -File Build-LinuxInstaller.ps1 -Version 1.0.0

# 输出: GOWordAgent-Install-linux-x64-1.0.0.run
```

### 构建文件说明

| 文件 | 用途 |
|------|------|
| `build-release.sh` | 构建发布版本（后端+插件） |
| `create-oneclick-installer.sh` | Linux 环境构建单文件安装器 |
| `Build-LinuxInstaller.ps1` | Windows PowerShell 构建脚本 |
| `installer-template.sh` | 安装器脚本模板 |

---

## 安装器技术细节

### 自解压原理

安装器是一个 shell 脚本，包含：

1. **脚本部分**：安装逻辑（前 N 字节）
2. **数据部分**：tar.gz 资源文件（附加在脚本后）

运行时使用 `tail` 提取数据部分：

```bash
SCRIPT_SIZE=12345  # 脚本大小
tail -c +$((SCRIPT_SIZE + 1)) "$0" > resources.tar.gz
tar -xzf resources.tar.gz
```

### 安装路径

```
~/.local/opt/gowordagent/          # 后端服务
~/.config/gowordagent/             # 配置文件
~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/  # WPS 插件
```

### 服务管理

安装器会尝试以下方式启动服务：

1. **Systemd**（优先）：创建用户服务 `gowordagent`
2. **手动脚本**：创建 `run.sh` 脚本管理进程

```bash
# Systemd 管理
systemctl --user status gowordagent
systemctl --user start gowordagent

# 手动脚本管理
~/.local/opt/gowordagent/run.sh status
~/.local/opt/gowordagent/run.sh start
```

---

## 故障排除

### 安装器无法运行

```bash
# 检查文件完整性
file GOWordAgent-Install-linux-x64-1.0.0.run

# 检查执行权限
ls -la GOWordAgent-Install-linux-x64-1.0.0.run

# 手动添加权限
chmod +x GOWordAgent-Install-linux-x64-1.0.0.run
```

### 服务启动失败

```bash
# 检查日志
tail -f /tmp/gowordagent-$USER.log

# 检查端口文件
cat /tmp/gowordagent-port-$USER.json

# 手动启动测试
~/.local/opt/gowordagent/gowordagent-server
```

### 依赖缺失

```bash
# 安装 tar（银河麒麟/Ubuntu）
sudo apt-get install tar

# 安装 zenity（可选，用于图形界面）
sudo apt-get install zenity
```

---

## 卸载

```bash
# 停止服务
systemctl --user stop gowordagent 2>/dev/null || \
    ~/.local/opt/gowordagent/run.sh stop

# 删除文件
rm -rf ~/.local/opt/gowordagent
rm -rf ~/.config/gowordagent
rm -rf ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
rm -f ~/.config/systemd/user/gowordagent.service
```

---

## 发布清单

向用户提供的文件：

```
发布包/
├── GOWordAgent-Install-linux-x64-1.0.0.run   # 单文件安装器
├── README.md                                  # 使用说明
└── INSTALL.md                                 # 安装指南
```

---

**创建日期**: 2026-04-13  
**版本**: 1.0.0
