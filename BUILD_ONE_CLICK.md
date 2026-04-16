# GOWordAgent Linux 一键安装包构建指南

## 📦 单文件安装包

### 方案一：一键安装脚本 + 资源目录（推荐）

#### 构建步骤

```bash
# 1. 先构建发布版本
./build-release.sh

# 2. 创建一键安装包目录
mkdir -p GOWordAgent-Installer
cp GOWordAgent-OneClick-Install.sh GOWordAgent-Installer/
cp -r release/gowordagent-linux-x64-1.0.0/backend GOWordAgent-Installer/
cp -r release/gowordagent-linux-x64-1.0.0/addon GOWordAgent-Installer/

# 3. 打包
chmod +x GOWordAgent-Installer/GOWordAgent-OneClick-Install.sh
tar -czf GOWordAgent-Installer-linux-x64-1.0.0.tar.gz GOWordAgent-Installer/
```

#### 用户使用

```bash
# 下载解压后
cd GOWordAgent-Installer
./GOWordAgent-OneClick-Install.sh
```

### 方案二：图形界面安装程序（高级）

使用 AppImage 格式，创建真正的单文件可执行程序。

#### 要求

需要在 Linux 环境下构建：

```bash
# 下载 appimagetool
wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage
chmod +x appimagetool-x86_64.AppImage

# 构建 AppImage
./build-appimage.sh
```

#### 输出

- `GOWordAgent-1.0.0-x86_64.AppImage` - 可双击运行的单文件

---

## 🚀 最简单的用户使用方式

### 方法 1：脚本安装（推荐）

```bash
# 下载
curl -O https://your-server/GOWordAgent-Installer-linux-x64-1.0.0.tar.gz

# 解压
tar -xzf GOWordAgent-Installer-linux-x64-1.0.0.tar.gz
cd GOWordAgent-Installer

# 运行安装（支持图形界面和命令行）
./GOWordAgent-OneClick-Install.sh
```

### 方法 2：桌面快捷方式

创建桌面文件 `GOWordAgent-Install.desktop`：

```ini
[Desktop Entry]
Name=安装 GOWordAgent
Comment=安装 GOWordAgent 智能校对插件
Exec=/path/to/GOWordAgent-OneClick-Install.sh
Type=Application
Terminal=true
Icon=application-x-executable
Categories=Office;
```

用户双击即可安装。

---

## 📋 安装包内容

```
GOWordAgent-Installer/
├── GOWordAgent-OneClick-Install.sh  # 安装程序（带图形界面支持）
├── backend/                          # 后端服务文件
│   ├── gowordagent-server
│   └── ...
└── addon/                            # WPS 插件文件
    ├── index.html
    ├── main.js
    └── ...
```

---

## 🎨 图形界面支持

安装脚本会自动检测并使用 `zenity` 提供图形界面：

```bash
# 安装 zenity（银河麒麟/Ubuntu）
sudo apt install zenity

# 安装 zenity（CentOS/RHEL）
sudo yum install zenity
```

如果系统没有 zenity，会自动回退到命令行界面。

---

## 🎯 一键安装脚本特性

| 特性 | 说明 |
|------|------|
| **图形界面** | 支持 zenity 图形对话框 |
| **命令行界面** | 无图形环境时自动回退 |
| **自动检测** | 检测 Systemd/手动服务管理 |
| **进度显示** | 5步安装流程，清晰明了 |
| **错误处理** | 完善的错误提示和处理 |
| **服务验证** | 安装后自动验证服务状态 |

---

## 📦 预编译安装包发布

### 文件命名

```
GOWordAgent-Installer-linux-x64-{VERSION}.tar.gz
```

### 发布文件清单

```
发布包/
├── GOWordAgent-Installer-linux-x64-1.0.0.tar.gz
├── README.md
└── INSTALL.md
```

### 用户文档

**README.md**：

```markdown
# GOWordAgent Linux 安装包

## 快速安装

```bash
tar -xzf GOWordAgent-Installer-linux-x64-1.0.0.tar.gz
cd GOWordAgent-Installer
./GOWordAgent-OneClick-Install.sh
```

## 使用说明

1. 运行安装脚本
2. 重启 WPS 文字
3. 在右侧边栏找到"智能校对"
4. 配置 API Key 并开始使用
```

---

## 🔧 高级定制

### 自定义安装路径

```bash
# 设置环境变量后运行安装
export INSTALL_DIR=/opt/gowordagent
./GOWordAgent-OneClick-Install.sh
```

### 静默安装

修改安装脚本，添加 `--silent` 参数支持（需要定制开发）。

---

## 📊 安装流程

```
┌─────────────────────────────────────────┐
│  1. 显示欢迎界面                        │
│     检查系统架构                         │
├─────────────────────────────────────────┤
│  2. 配置安装路径                         │
│     确认安装                             │
├─────────────────────────────────────────┤
│  3. 解压安装文件                         │
│     复制后端和插件                       │
├─────────────────────────────────────────┤
│  4. 配置系统服务                         │
│     Systemd 或手动脚本                   │
├─────────────────────────────────────────┤
│  5. 验证安装                             │
│     检查端口文件                         │
│     健康检查                             │
├─────────────────────────────────────────┤
│  显示安装完成界面                        │
│  提供服务管理命令                        │
└─────────────────────────────────────────┘
```

---

## ✅ 测试检查清单

- [ ] 在银河麒麟 V10 上测试安装
- [ ] 在 Ubuntu 上测试安装
- [ ] 测试图形界面（有 zenity）
- [ ] 测试命令行界面（无 zenity）
- [ ] 测试 Systemd 服务管理
- [ ] 测试手动服务管理
- [ ] 测试卸载功能
- [ ] 测试升级安装

---

**创建日期**: 2026-04-13  
**版本**: 1.0.0
