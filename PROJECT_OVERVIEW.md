# GOWordAgent 项目介绍

> **版本**：银河麒麟 V10 改造版 v1.0  
> **日期**：2026-04-13

---

## 项目概述

GOWordAgent 是一个跨平台的 Word/WPS 智能校对插件，支持 AI 驱动的文本校对功能。

### 双版本架构

| 版本 | 目标平台 | 技术栈 | 状态 |
|------|----------|--------|------|
| **Word 版** | Windows + Microsoft Word | VSTO + WPF | 保持稳定 |
| **WPS 版** | 银河麒麟 V10 + WPS Office | .NET 8 + HTML/JS | 新增开发 |

---

## 项目结构

```
GOWordAgent/                                    # 项目根目录
│
├── GOWordAgentAddIn/                           # 【Word VSTO 版 - 保持原有】
│   ├── gowordagent.csproj                      # VSTO 项目文件
│   ├── ThisAddIn.cs                            # 插件入口类
│   ├── GOWordAgentPaneWpf.xaml/.cs             # WPF 侧边栏（1200+ 行）
│   ├── WordDocumentService.cs                  # Word COM 操作封装
│   ├── WordProofreadController.cs              # 校对控制逻辑
│   ├── ProofreadService.cs                     # 校对服务（将被 Core 替代）
│   ├── BaseLLMService.cs                       # LLM 基类（将被 Core 替代）
│   └── ...                                     # 其他原有文件
│
├── GOWordAgent.Core/                           # 【新增 - .NET 8 共享类库】
│   ├── GOWordAgent.Core.csproj                 # Core 项目文件
│   ├── Models/
│   │   └── ProofreadModels.cs                  # 数据模型定义
│   ├── Config/
│   │   └── ConfigManager.cs                    # 跨平台配置管理
│   └── Services/
│       ├── ILLMService.cs                      # LLM 服务接口
│       ├── BaseLLMService.cs                   # LLM 抽象基类
│       ├── DeepSeekService.cs                  # DeepSeek API 适配
│       ├── GLMService.cs                       # 智谱 GLM API 适配
│       ├── OllamaService.cs                    # Ollama 本地模型适配
│       ├── LLMServiceFactory.cs                # 服务工厂
│       ├── LLMServiceException.cs              # 异常定义
│       ├── ProofreadService.cs                 # 校对服务核心
│       ├── DocumentSegmenter.cs                # 智能文档分段
│       ├── ProofreadCacheManager.cs            # LRU 缓存管理
│       └── ProofreadIssueParser.cs             # 校对结果解析
│
├── GOWordAgent.WpsService/                     # 【新增 - .NET 8 后端服务】
│   ├── GOWordAgent.WpsService.csproj           # 服务项目文件
│   ├── Program.cs                              # 服务入口（Kestrel 配置）
│   └── Controllers/
│       └── ProofreadController.cs              # API 控制器（校对/配置/健康）
│
├── GOWordAgent.WpsAddon/                       # 【新增 - WPS 加载项】
│   ├── package.json                            # wpsjs 加载项配置
│   ├── index.html                              # 侧边栏主页面
│   ├── main.js                                 # 加载项入口
│   ├── css/
│   │   └── style.css                           # 侧边栏样式
│   └── js/
│       ├── apiClient.js                        # 后端 HTTP 通信
│       ├── documentService.js                  # WPS JS API 封装
│       ├── proofreadService.js                 # 校对工作流
│       └── uiController.js                     # UI 控制逻辑
│
├── Scripts/                                    # 【新增 - 部署脚本】
│   ├── install.sh                              # 一键安装脚本
│   ├── uninstall.sh                            # 卸载脚本
│   └── gowordagent.service                     # systemd 服务配置
│
└── Docs/                                       # 文档目录
    ├── ARCHITECTURE.md                         # 架构说明
    ├── API.md                                  # API 文档
    ├── KYLIN_V10_MIGRATION_RECORD.md           # 改造记录
    ├── CODE_OPTIMIZATION_PLAN.md               # 优化方案
    └── TEST_GUIDE.md                           # 测试指南
```

---

## 各模块详细介绍

### 1. GOWordAgentAddIn（Word VSTO 版）

**用途**：原有 Windows Word 插件，保持不变继续维护

**核心文件**：

| 文件 | 作用 | 说明 |
|------|------|------|
| `ThisAddIn.cs` | 插件生命周期管理 | VSTO 入口，处理启动/关闭 |
| `GOWordAgentPaneWpf.xaml` | WPF 侧边栏 UI | XAML 定义的界面布局 |
| `GOWordAgentPaneWpf.xaml.cs` | 侧边栏逻辑 | 1200+ 行，处理所有 UI 交互 |
| `WordDocumentService.cs` | Word COM 封装 | 安全的 COM 对象管理 |
| `WordProofreadController.cs` | 校对控制 | 协调校对流程和文档操作 |
| `ProofreadResultRenderer.cs` | 结果渲染 | 生成校对报告 UI |

**维护策略**：
- Bug 修复继续在原项目进行
- 新功能优先在 Core 实现，视情况同步到 VSTO

---

### 2. GOWordAgent.Core（共享类库）

**用途**：.NET 8 类库，被 WPS 后端服务引用，也可被其他 .NET 项目使用

**核心设计原则**：
- 无 UI 依赖（移除 WPF/WinForms）
- 跨平台（Windows/Linux/macOS）
- 线程安全

#### 2.1 Models（数据模型）

**文件**：`ProofreadModels.cs`

**定义**：
```csharp
ParagraphResult     // 段落校对结果
ProofreadProgressArgs   // 进度事件参数
ProofreadIssueItem      // 单个校对问题
ProviderItem            // AI 提供商
AIProvider              // 提供商枚举
```

**用途**：前后端数据交换的标准格式

#### 2.2 Config（配置管理）

**文件**：`ConfigManager.cs`

**功能**：
- Windows：DPAPI 加密 + HMAC 完整性校验
- Linux：AES-GCM + /etc/machine-id 派生密钥
- XDG 路径规范（~/.config/gowordagent/）

**关键方法**：
```csharp
LoadConfig()           // 加载配置（自动识别平台）
SaveConfig()           // 保存配置（自动选择加密方式）
GetProofreadPromptForMode()  // 获取校验模式提示词
```

#### 2.3 Services（核心服务）

##### ILLMService / BaseLLMService

**功能**：LLM 服务抽象层

**支持的提供商**：
- DeepSeek（deepseek-chat）
- 智谱 GLM（glm-4.7）
- Ollama 本地（llama2/qwen 等）

**关键方法**：
```csharp
SendMessageAsync()           // 普通对话
SendMessagesWithHistoryAsync()   // 带历史记录的对话
SendProofreadMessageAsync()      // 校对专用（更低 temperature）
```

##### ProofreadService

**功能**：校对业务流程 orchestration

**特性**：
- 并发控制（SemaphoreSlim，默认 5 并发）
- LRU 缓存（基于 SHA256 内容哈希）
- 进度事件（支持 SSE/轮询）
- 取消令牌（支持中途取消）

**工作流程**：
```
1. 文档分段（DocumentSegmenter）
2. 并行校对（带缓存检查）
3. 结果解析（ProofreadIssueParser）
4. 生成报告
```

##### DocumentSegmenter

**功能**：智能文档分段

**策略**：
- 目标段大小：1500 字符
- 重叠区域：100 字符（防止边界遗漏）
- 分割点：句号 > 分号 > 逗号

##### ProofreadCacheManager

**功能**：内存级 LRU 缓存

**特性**：
- 最大条目：1000
- 线程安全（lock + ConcurrentDictionary）
- 自动淘汰（按访问频率）
- 缓存统计（条目数/内存占用）

##### ProofreadIssueParser

**功能**：解析 AI 返回的校对结果

**支持的格式**：
```
【第1处】类型：错别字｜严重度：high
原文：错误的词
修改：正确的词
理由：解释说明
```

**容错处理**：
- 正则超时保护（5 秒）
- 多格式兼容（宽松匹配）

---

### 3. GOWordAgent.WpsService（后端服务）

**用途**：.NET 8 Minimal API，作为 WPS 加载项的后端

**技术选型**：
- Kestrel（内置 Web 服务器）
- 自动端口分配（0 = 随机端口）
- CORS（允许 WPS WebView 访问）
- Self-Contained 部署

#### Program.cs

**功能**：服务入口

**关键逻辑**：
```csharp
1. 配置 Kestrel（自动端口）
2. 注册依赖注入（ILLMService）
3. 加载配置（ConfigManager.LoadConfig）
4. 启动后写入端口文件（/tmp/gowordagent-port.json）
5. 退出时清理端口文件
```

#### ProofreadController.cs

**API 端点**：

| 方法 | 路径 | 功能 |
|------|------|------|
| GET | `/api/proofread/health` | 健康检查 |
| POST | `/api/proofread` | 执行校对 |
| GET | `/api/proofread/config` | 获取配置 |
| POST | `/api/proofread/config` | 保存配置 |

**关键转换**：后端返回的结果包含**字符偏移量**（替代 Word Range 对象）

```csharp
// 后端计算全局偏移量
var startOffset = para.StartOffset + localOffset;
var endOffset = startOffset + item.Original.Length;

// 前端直接使用
range = doc.Range(startOffset, endOffset);
```

---

### 4. GOWordAgent.WpsAddon（WPS 加载项）

**用途**：WPS Office 插件，HTML/JS/CSS 实现

**技术约束**：
- 纯原生 JS（无 Vue/React，避免构建链）
- XMLHttpRequest（不用 fetch，兼容性最好）
- ES5/ES6 子集（避免高级语法糖）

#### index.html

**结构**：
```
状态栏（连接状态）
├── 设置面板（AI 配置、校验模式）
├── 结果面板（问题列表、进度）
└── 操作栏（开始校对、清空）
```

#### js/apiClient.js

**功能**：后端通信封装

**核心方法**：
```javascript
discoverService()     // 从 /tmp/gowordagent-port.json 读取端口
get() / post()        // XMLHttpRequest 封装
healthCheck()         // 健康检查
proofread()           // 执行校对
```

#### js/documentService.js

**功能**：WPS JS API 封装

**与 Word VSTO 的区别**：

| 功能 | Word VSTO | WPS JS |
|------|-----------|--------|
| 获取文本 | `doc.Content.Text` | `doc.Paragraphs.Item(i).Range.Text` |
| 定位 | `Range.Select()` | `doc.Range(start, end).Select()` |
| 修订 | `doc.TrackRevisions = true` | `doc.TrackRevisions = true` |
| 批注 | `doc.Comments.Add(range, text)` | `doc.Comments.Add(range, text)` |

**关键创新**：使用**字符偏移量**定位（替代 Find.Execute）

```javascript
// 获取文档时记录偏移量
for (var i = 1; i <= doc.Paragraphs.Count; i++) {
    paragraphs.push({
        index: i - 1,
        start: offset,           // 字符偏移量
        end: offset + text.length,
        text: text
    });
    offset += text.length;
}

// 应用修订时使用偏移量
var range = doc.Range(startOffset, endOffset);
range.Delete();
range.InsertAfter(replacement);
```

#### js/proofreadService.js

**功能**：校对工作流编排

**流程**：
```javascript
1. 发现后端服务（轮询 /tmp/gowordagent-port.json）
2. 提取文档文本（含偏移量）
3. 发送校对请求
4. 接收结果
5. 应用修订（按偏移量倒序处理）
6. 显示问题列表
```

#### js/uiController.js

**功能**：UI 状态管理

**事件处理**：
- 连接按钮点击
- 校验模式切换
- 开始校对
- 问题项点击定位

---

### 5. Scripts（部署脚本）

#### install.sh

**功能**：一键安装

**步骤**：
```bash
1. 检查架构（x86_64）
2. 检查 WPS
3. 复制后端到 /opt/gowordagent/
4. 注册 systemd 用户服务
5. 安装 WPS 加载项到 ~/.local/share/Kingsoft/wps/jsaddons/
6. 启动服务
```

#### uninstall.sh

**功能**：卸载

**清理内容**：
- 停止并禁用 systemd 服务
- 删除 /opt/gowordagent/
- 删除 WPS 加载项
- 可选删除配置

#### gowordagent.service

**功能**：systemd 用户服务配置

**特性**：
- 开机自启（user 级别）
- 崩溃自动重启（Restart=on-failure）
- 日志输出到 journal

---

## 数据流向图

### Word 版（VSTO）

```
用户点击"开始校对"
    ↓
GOWordAgentPaneWpf.xaml.cs
    ↓
WordProofreadController
    ↓
ProofreadService（本进程）
    ↓
BaseLLMService → HTTP → AI API
    ↓
结果写回 Word（TrackRevisions）
```

### WPS 版（前后端分离）

```
用户点击"开始校对"
    ↓
index.html / uiController.js
    ↓
proofreadService.js
    ↓
apiClient.js → HTTP → WpsService（.NET 8）
    ↓
ProofreadController → ProofreadService
    ↓
BaseLLMService → HTTP → AI API
    ↓
结果返回（JSON + 偏移量）
    ↓
documentService.js → WPS JS API
    ↓
结果写回 WPS（TrackRevisions）
```

---

## 关键技术决策

| 决策 | 选择 | 理由 |
|------|------|------|
| 前后端分离 | HTTP API | WPS WebView 限制，无法直接调用 .NET |
| 定位方式 | 字符偏移量 | 避免 Find.Execute 重复匹配问题 |
| 前端框架 | 原生 HTML/JS | 避免构建链，兼容性最好 |
| JSON 库 | Newtonsoft.Json | 保留现有代码，不迁移风险 |
| 部署方式 | Self-Contained | 零运行时依赖 |
| 加密方案 | AES-GCM + machine-id | Linux 原生兼容 |

---

## 开发/调试指南

### 开发环境

**Windows（推荐）**：
- Visual Studio 2022 / VS Code
- .NET 8 SDK
- WPS Office Windows 版（JS API 调试）

**银河麒麟 V10**：
- 直接部署测试
- 使用 journalctl 查看日志

### 调试技巧

**后端调试**：
```bash
# 手动运行查看输出
/opt/gowordagent/gowordagent-server

# 查看日志
journalctl --user -u gowordagent -f
```

**前端调试**：
```javascript
// 在 WPS 加载项中打开控制台（如果有）
// 或使用 alert() 输出调试信息
alert('调试信息: ' + JSON.stringify(data));
```

**API 测试**：
```bash
# 健康检查
curl http://127.0.0.1:$PORT/api/proofread/health

# 校对接口
curl -X POST http://127.0.0.1:$PORT/api/proofread \
  -H "Content-Type: application/json" \
  -d '{"text":"测试","paragraphs":[],"provider":"DeepSeek","apiKey":"xxx"}'
```

---

## 维护注意事项

### 代码同步

当修改以下文件时，需要同步到两个版本：

| 文件 | Word VSTO | WPS Core | 同步策略 |
|------|-----------|----------|----------|
| `ProofreadService.cs` | ✅ | ✅ | 优先改 Core，视情况同步 |
| `BaseLLMService.cs` | ✅ | ✅ | 优先改 Core，视情况同步 |
| `ProofreadIssueParser.cs` | ✅ | ✅ | 优先改 Core，视情况同步 |
| `ConfigManager.cs` | ✅（DPAPI） | ✅（AES-GCM） | 平台相关，分别维护 |
| UI 相关 | ✅（WPF） | ✅（HTML/JS） | 独立维护 |

### 版本发布

**Word VSTO 版**：
- 更新版本号
- 生成安装包（ClickOnce/MSI）
- 发布到内部更新服务器

**WPS 版**：
- 更新 `package.json` 版本
- 构建后端（linux-x64 Self-Contained）
- 打包发布包（backend/ + addon/ + Scripts/）
- 提供 install.sh

---

*文档结束*
