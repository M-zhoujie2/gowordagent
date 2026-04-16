# 快速构建指南

## 在 WSL (Windows) 上构建

```powershell
# 步骤 1: 打开 WSL
wsl

# 步骤 2: 在 WSL 中构建
cd /mnt/d/Project/word/GOWordAgentAddIn-master
./build-release.sh

# 步骤 3: 构建单文件安装器
./create-oneclick-installer.sh

# 退出 WSL
exit

# 步骤 4: (可选) 在 PowerShell 中构建安装器
powershell -ExecutionPolicy Bypass -File Build-LinuxInstaller.ps1 -Version 1.0.0
```

## 在 Linux 上构建

```bash
# 步骤 1: 构建发布版本
./build-release.sh

# 步骤 2: 构建单文件安装器
./create-oneclick-installer.sh
```

## 构建输出

构建完成后，你会得到：

| 文件 | 说明 | 用途 |
|------|------|------|
| `release/gowordagent-linux-x64-1.0.0/` | 标准发布包 | 手动部署 |
| `GOWordAgent-Install-linux-x64-1.0.0.run` | 单文件安装器 | 双击安装 |

## 测试安装器

```bash
# 在 WSL 或 Linux 中测试
chmod +x GOWordAgent-Install-linux-x64-1.0.0.run
./GOWordAgent-Install-linux-x64-1.0.0.run
```

## 发布给用户

复制以下文件：
- `GOWordAgent-Install-linux-x64-1.0.0.run` (安装器)
- `Docs/INSTALLER_README.md` (用户说明)

或者只给安装器文件，用户双击即可安装。
