# GOWordAgent 智能校对

银河麒麟 V10 版本 - WPS Office 智能校对插件

## 项目简介

GOWordAgent 是一个基于 **.NET 8** 开发的 **WPS Office for Linux** 插件，提供 AI 智能校对功能。它将文档内容发送给 AI 模型，获取修改建议，并以修订和批注的形式写回 WPS 文档。

## 功能特性

### 1. 智能校对
- **精准校验**：针对错别字、语病、术语不一致等精确纠错
- **全文校验**：系统性校对，包括语法、标点、用词、逻辑连贯性
- 自定义提示词，支持两种校验模式独立配置

### 2. 修订与批注
- 使用 WPS 原生修订功能（TrackRevisions）
- 标记原文（删除线）和建议文本（下划线）
- 添加批注说明错误类型和修改理由
- 通过 WPS 审阅面板接受或拒绝修订

### 3. 问题定位
- 聊天框显示发现的每个问题
- 点击"定位"按钮跳转到文档对应位置
- 支持严重程度标记（high/medium/low）

### 4. 多 AI 提供商支持
- DeepSeek
- 智谱 AI (GLM)
- Ollama 本地模型

## 技术架构

### 系统架构

```
银河麒麟 V10 桌面
│
├── WPS 文字 (Linux)
│   └── WPS 加载项 (HTML/JS/CSS)
│       ├── 设置面板
│       ├── 校对结果列表  
│       └── HTTP 通信
│
└── .NET 8 后端服务 (Self-Contained)
    ├── Minimal API (Kestrel)
    ├── GOWordAgent.Core (核心逻辑)
    │   ├── ProofreadService (校对服务)
    │   ├── BaseLLMService (LLM 基类)
    │   ├── DocumentSegmenter (文档分段)
    │   └── Cache/Parser/Models
    └── 配置管理 (AES-GCM 加密)
```

### 项目结构

```
GOWordAgent/
├── GOWordAgent.Core/              # .NET 8 共享类库
│   ├── Config/
│   │   └── ConfigManager.cs       # 跨平台配置管理 (AES-GCM)
│   ├── Models/
│   │   └── ProofreadModels.cs     # 数据模型
│   └── Services/
│       ├── ProofreadService.cs    # 校对服务核心
│       ├── BaseLLMService.cs      # LLM 服务基类
│       ├── DocumentSegmenter.cs   # 智能文档分段
│       ├── ProofreadCacheManager.cs # LRU 缓存
│       └── ProofreadIssueParser.cs  # 结果解析
│
├── GOWordAgent.WpsService/        # .NET 8 后端服务
│   ├── Program.cs                 # 服务入口
│   └── Controllers/
│       └── ProofreadController.cs # API 控制器
│
├── GOWordAgent.WpsAddon/          # WPS 加载项
│   ├── index.html                 # 主页面
│   ├── main.js                    # 入口脚本
│   ├── css/style.css              # 样式
│   └── js/
│       ├── apiClient.js           # 后端通信
│       ├── documentService.js     # WPS JS API 封装
│       ├── proofreadService.js    # 校对工作流
│       └── uiController.js        # UI 控制
│
├── Scripts/                       # 部署脚本
│   ├── install.sh                 # 安装脚本
│   ├── uninstall.sh               # 卸载脚本
│   └── gowordagent.service        # systemd 配置
│
└── Docs/                          # 文档
    ├── PROJECT_OVERVIEW.md        # 项目介绍
    ├── FILE_REFERENCE.md          # 文件速查
    ├── KYLIN_V10_BUILD.md         # 构建指南
    └── TEST_GUIDE.md              # 测试指南
```

## 开发环境

### 必需软件
- **.NET 8 SDK** (https://dotnet.microsoft.com/download/dotnet/8.0)
- **WPS Office for Linux** 12.1.2.25838+
- **银河麒麟 V10 SP1** (x86_64)

### 可选软件
- Visual Studio 2022 或 VS Code（Windows 开发）
- Postman 或 curl（API 测试）

## 快速开始

### 方式一：单文件安装器（推荐）

```bash
# 下载单文件安装器
wget https://your-server/GOWordAgent-Install-linux-x64-1.0.0.run

# 添加执行权限并运行
chmod +x GOWordAgent-Install-linux-x64-1.0.0.run
./GOWordAgent-Install-linux-x64-1.0.0.run
```

支持图形界面（双击安装）和命令行界面，自动配置 Systemd 服务。

### 方式二：脚本部署

```bash
# 下载预编译发布包
tar -xzf gowordagent-linux-x64-1.0.0.tar.gz
cd gowordagent-linux-x64-1.0.0

# 一键部署
./deploy-linux.sh
```

详见 [Docs/ONE_CLICK_INSTALLER.md](Docs/ONE_CLICK_INSTALLER.md)

### 方式二：手动构建部署

#### 1. 构建项目

```bash
# 克隆仓库
git clone <repository-url>
cd GOWordAgent

# 构建 Core 类库
dotnet build GOWordAgent.Core/GOWordAgent.Core.csproj -c Release

# 构建后端服务（linux-x64）
dotnet publish GOWordAgent.WpsService/GOWordAgent.WpsService.csproj \
    -c Release -r linux-x64 --self-contained true \
    -p:PublishSingleFile=true -o ./publish/backend
```

### 2. 打包

```bash
mkdir -p release/backend release/addon
cp -r publish/backend/* release/backend/
cp -r GOWordAgent.WpsAddon/* release/addon/
cp Scripts/* release/
```

### 3. 部署到银河麒麟 V10

```bash
# 复制到目标机器
scp -r release/ user@kylin-host:/tmp/gowordagent-release

# SSH 登录并安装
ssh user@kylin-host
cd /tmp/gowordagent-release
./install.sh
```

### 4. 验证

```bash
# 检查服务状态
systemctl --user status gowordagent

# 测试 API
PORT=$(cat /tmp/gowordagent-port.json | grep -o '"port":[0-9]*' | cut -d: -f2)
curl http://127.0.0.1:$PORT/api/proofread/health
```

## 配置说明

配置文件位置：`~/.config/gowordagent/config.dat`

配置内容包括：
- AI 提供商和 API 配置
- 自定义提示词（支持精准/全文两种模式）
- 自动连接设置

**加密方式**：使用 AES-GCM + `/etc/machine-id` 派生密钥

## 卸载

```bash
./Scripts/uninstall.sh
```

## 文档

| 文档 | 说明 |
|------|------|
| [PROJECT_OVERVIEW.md](Docs/PROJECT_OVERVIEW.md) | 项目全景介绍 |
| [FILE_REFERENCE.md](Docs/FILE_REFERENCE.md) | 文件速查手册 |
| [KYLIN_V10_BUILD.md](KYLIN_V10_BUILD.md) | 构建和部署指南 |
| [TEST_GUIDE.md](TEST_GUIDE.md) | 测试指南 |
| [MIGRATION_COMPLETE.md](MIGRATION_COMPLETE.md) | 改造完成报告 |

## 技术特点

- **跨平台**：.NET 8 Self-Contained 部署，零运行时依赖
- **前后分离**：WPS 加载项（HTML/JS）+ .NET 后端（HTTP API）
- **精确定位**：使用字符偏移量替代查找，避免重复匹配
- **安全加密**：AES-GCM 加密配置，基于机器 ID 派生密钥
- **并发处理**：支持 3-5 段并行校对
- **智能缓存**：内存级 LRU 缓存，基于内容哈希

## 系统要求

- **操作系统**：银河麒麟 V10 SP1 (x86_64)
- **WPS Office**：12.1.2.25838 或更高版本
- **运行内存**：4GB+
- **磁盘空间**：200MB（后端服务）

## 许可证

MIT License

---

**注意**：本项目专为银河麒麟 V10 + WPS Office 环境开发，不支持 Windows/Microsoft Word。
