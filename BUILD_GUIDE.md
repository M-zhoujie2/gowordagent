# GOWordAgent 构建指南

## 开发者构建流程

### 前置要求

- **Windows 开发**: WSL (Windows Subsystem for Linux)
- **Linux 开发**: 任何 Linux 发行版
- **.NET 8 SDK**: https://dotnet.microsoft.com/download/dotnet/8.0

---

## 快速构建

### 在 Linux/WSL 上构建

```bash
# 1. 构建发布版本
./build-release.sh

# 2. (可选) 构建单文件安装器
./create-oneclick-installer.sh
```

### 在 Windows 上构建（需要 WSL）

```powershell
# 1. 在 WSL 中构建发布版本
wsl ./build-release.sh

# 2. 构建单文件安装器
powershell -ExecutionPolicy Bypass -File Build-LinuxInstaller.ps1 -Version 1.0.0
```

---

## 构建输出

```
.
├── release/
│   └── gowordagent-linux-x64-1.0.0/    # 标准发布包
│       ├── backend/                     # 后端服务
│       ├── addon/                       # WPS 插件
│       └── deploy-linux.sh              # 部署脚本
│
└── GOWordAgent-Install-linux-x64-1.0.0.run  # 单文件安装器
```

---

## 构建脚本说明

| 脚本 | 用途 | 运行环境 |
|------|------|----------|
| `build-release.sh` | 构建发布版本 | Linux/WSL |
| `create-oneclick-installer.sh` | 构建单文件安装器 | Linux/WSL |
| `Build-LinuxInstaller.ps1` | PowerShell 构建脚本 | Windows (需WSL) |

---

## 发布给用户

### 推荐方式：单文件安装器

向用户提供：
- `GOWordAgent-Install-linux-x64-1.0.0.run`
- `Docs/INSTALLER_README.md` (使用说明)

### 备选方式：标准发布包

向用户提供：
- `gowordagent-linux-x64-1.0.0.tar.gz`
- 解压后运行 `./deploy-linux.sh`

---

## 故障排除

### WSL 不可用

如果没有 WSL，可以：
1. 在虚拟机中构建
2. 使用 GitHub Actions 等 CI/CD 服务
3. 在实体 Linux 机器上构建

### 构建失败

```bash
# 检查 .NET SDK
dotnet --version

# 检查运行时
wsl dotnet --version

# 清理并重建
rm -rf release/
./build-release.sh
```

---

## 文档索引

| 文档 | 说明 |
|------|------|
| [Docs/ONE_CLICK_INSTALLER.md](Docs/ONE_CLICK_INSTALLER.md) | 单文件安装器详细说明 |
| [Docs/INSTALLER_README.md](Docs/INSTALLER_README.md) | 用户安装说明 |
| [Docs/PROJECT_OVERVIEW.md](Docs/PROJECT_OVERVIEW.md) | 项目介绍 |
| [Docs/FILE_REFERENCE.md](Docs/FILE_REFERENCE.md) | 文件速查手册 |
