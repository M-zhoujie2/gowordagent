# GOWordAgent 文件速查手册

快速查找每个文件的作用和关键信息。

---

## 符号说明

| 标记 | 含义 |
|------|------|
| 🟢 | 新增文件（麒麟改造） |
| 🟡 | 修改文件（适配改造） |
| ⚪ | 原有文件（保持不变） |
| ⭐ | 核心文件（重点关注） |

---

## GOWordAgent.Core（.NET 8 类库）

| 文件 | 标记 | 用途 | 关键类/方法 |
|------|------|------|-------------|
| `GOWordAgent.Core.csproj` | 🟢 | 项目配置 | TargetFramework: net8.0 |
| **Models/ProofreadModels.cs** | 🟢 | 数据模型 | `ParagraphResult`, `ProofreadIssueItem` |
| **Config/ConfigManager.cs** | 🟢⭐ | 配置管理 | `LoadConfig()`, `SaveConfig()`, Linux 加密 |
| Services/ILLMService.cs | 🟢 | 接口定义 | `ILLMService` 接口 |
| **Services/BaseLLMService.cs** | 🟢⭐ | LLM 基类 | `SendProofreadMessageAsync()` |
| Services/DeepSeekService.cs | 🟢 | DeepSeek 适配 | 继承 BaseLLMService |
| Services/GLMService.cs | 🟢 | GLM 适配 | 继承 BaseLLMService |
| Services/OllamaService.cs | 🟢 | Ollama 适配 | 本地模型支持 |
| Services/LLMServiceFactory.cs | 🟢 | 服务工厂 | `CreateService()` 工厂方法 |
| Services/LLMServiceException.cs | 🟢 | 异常定义 | `LLMServiceException` |
| **Services/ProofreadService.cs** | 🟢⭐ | 校对核心 | `ProofreadDocumentAsync()`, 并发控制 |
| Services/DocumentSegmenter.cs | 🟢 | 文档分段 | `SplitIntoParagraphs()`, 1500字/段 |
| Services/ProofreadCacheManager.cs | 🟢 | 缓存管理 | `TryGetCachedResult()`, LRU 缓存 |
| Services/ProofreadIssueParser.cs | 🟢 | 结果解析 | `ParseProofreadItems()`, 正则解析 |

---

## GOWordAgent.WpsService（.NET 8 后端）

| 文件 | 标记 | 用途 | 关键配置 |
|------|------|------|----------|
| `GOWordAgent.WpsService.csproj` | 🟢 | 项目配置 | SelfContained=true, linux-x64 |
| **Program.cs** | 🟢⭐ | 服务入口 | Kestrel 配置, 端口文件写入 |
| **Controllers/ProofreadController.cs** | 🟢⭐ | API 控制器 | `/api/proofread`, 偏移量计算 |

### API 端点速查

```
GET    /api/proofread/health     -> 健康检查
POST   /api/proofread            -> 执行校对
GET    /api/proofread/config     -> 获取配置
POST   /api/proofread/config     -> 保存配置
```

---

## GOWordAgent.WpsAddon（WPS 加载项）

| 文件 | 标记 | 用途 | 关键函数 |
|------|------|------|----------|
| `package.json` | 🟢 | 加载项配置 | wps.id: com.gowordagent.addin |
| **index.html** | 🟢⭐ | 主页面 | 侧边栏 UI 结构 |
| `main.js` | 🟢 | 入口脚本 | 初始化流程 |
| css/style.css | 🟢 | 样式表 | 面板/按钮/问题列表样式 |
| **js/apiClient.js** | 🟢⭐ | HTTP 通信 | `discoverService()`, `proofread()` |
| **js/documentService.js** | 🟢⭐ | 文档操作 | `getDocumentText()`, `applyAtOffset()` |
| **js/proofreadService.js** | 🟢⭐ | 工作流 | `startProofread()`, `applyRevisions()` |
| js/uiController.js | 🟢 | UI 控制 | 事件绑定, 状态更新 |

### JS API 速查

```javascript
// apiClient.js
ApiClient.discoverService()   // 发现后端端口
ApiClient.proofread(data, cb) // 执行校对

// documentService.js  
DocumentService.getDocumentText()              // 获取文档+偏移量
DocumentService.applyAtOffset(s, e, text, c)   // 应用修订
DocumentService.navigateToOffset(s, e)         // 跳转到位置

// proofreadService.js
ProofreadWorkflow.init()           // 初始化
ProofreadWorkflow.connect()        // 连接后端
ProofreadWorkflow.startProofread() // 开始校对
```

---

## GOWordAgentAddIn（Word VSTO - 原有）

| 文件 | 标记 | 用途 | 说明 |
|------|------|------|------|
| `gowordagent.csproj` | ⚪ | VSTO 项目 | 保持不变 |
| `ThisAddIn.cs` | ⚪ | 插件入口 | 保持不变 |
| `GOWordAgentPaneWpf.xaml` | ⚪ | WPF 界面 | 保持不变 |
| `GOWordAgentPaneWpf.xaml.cs` | ⚪ | WPF 逻辑 | 保持不变 |
| `WordDocumentService.cs` | ⚪ | COM 封装 | 保持不变 |
| `WordProofreadController.cs` | ⚪ | 校对控制 | 保持不变 |
| ProofreadService.cs | ⚪ | 校对服务 | 将被 Core 替代（保留兼容） |
| BaseLLMService.cs | ⚪ | LLM 基类 | 将被 Core 替代（保留兼容） |

---

## Scripts（部署脚本）

| 文件 | 标记 | 用途 | 关键操作 |
|------|------|------|----------|
| **install.sh** | 🟢⭐ | 安装脚本 | 复制文件+注册服务+启动 |
| **uninstall.sh** | 🟢 | 卸载脚本 | 停止服务+删除文件 |
| **gowordagent.service** | 🟢 | systemd 配置 | 用户级服务, 自动重启 |

### 安装路径

```
/opt/gowordagent/                           # 后端服务
├── gowordagent-server                      # 主程序
└── *.dll                                   # 依赖库

~/.local/share/Kingsoft/wps/jsaddons/       # WPS 加载项
└── com.gowordagent.addin/
    ├── index.html
    ├── main.js
    ├── css/
    └── js/

~/.config/gowordagent/                      # 配置
└── config.dat                              # AES 加密配置

/tmp/gowordagent-port.json                  # 运行时端口文件
```

---

## 文档文件

| 文件 | 标记 | 用途 | 阅读场景 |
|------|------|------|----------|
| `PROJECT_OVERVIEW.md` | 🟢 | 项目全景介绍 | 快速了解架构 |
| `FILE_REFERENCE.md` | 🟢 | 文件速查（本文件）| 查找文件作用 |
| `KYLIN_V10_BUILD.md` | 🟢 | 构建指南 | 如何编译发布 |
| `TEST_GUIDE.md` | 🟢 | 测试指南 | 如何测试验证 |
| `Docs/KYLIN_V10_MIGRATION_RECORD.md` | 🟢 | 改造记录 | 了解改造过程 |
| `Docs/CODE_OPTIMIZATION_PLAN.md` | 🟢 | 优化方案 | 未来优化方向 |
| `MIGRATION_COMPLETE.md` | 🟢 | 完成报告 | 改造总结 |

---

## 按角色查看

### 如果你是后端开发者

**必读文件**：
1. `GOWordAgent.Core/Services/ProofreadService.cs` - 了解校对流程
2. `GOWordAgent.WpsService/Controllers/ProofreadController.cs` - 了解 API 设计
3. `GOWordAgent.Core/Config/ConfigManager.cs` - 了解配置加密

**调试方法**：
```bash
dotnet run --project GOWordAgent.WpsService
curl http://localhost:PORT/api/proofread/health
```

### 如果你是前端开发者

**必读文件**：
1. `GOWordAgent.WpsAddon/js/documentService.js` - 了解 WPS API
2. `GOWordAgent.WpsAddon/js/apiClient.js` - 了解后端通信
3. `GOWordAgent.WpsAddon/index.html` - 了解 UI 结构

**调试方法**：
- 在 WPS 中使用 `alert()` 或 `console.log()`
- 使用 Python 测试服务器模拟后端

### 如果你是测试人员

**必读文件**：
1. `TEST_GUIDE.md` - 完整测试流程
2. `Scripts/install.sh` - 了解部署步骤

**测试要点**：
- Day 1 PoC：HTTP 通信测试（生死线）
- 功能测试：校对全流程
- 异常测试：后端断开、网络异常

### 如果你是运维人员

**必读文件**：
1. `Scripts/install.sh` - 安装流程
2. `Scripts/uninstall.sh` - 卸载流程
3. `Scripts/gowordagent.service` - 服务配置

**运维命令**：
```bash
# 查看服务状态
systemctl --user status gowordagent

# 查看日志
journalctl --user -u gowordagent -f

# 手动重启
systemctl --user restart gowordagent
```

---

## 修改影响范围

### 修改 Core 文件（影响两个版本）

```
修改 ProofreadService.cs
    ↓
Word VSTO 版：需决定是否同步更新
WPS 版：自动生效（重新编译服务）
```

### 修改 WpsService 文件（仅影响 WPS 版）

```
修改 ProofreadController.cs
    ↓
仅 WPS 版需要重新部署
```

### 修改 WpsAddon 文件（仅影响 WPS 版）

```
修改 js/documentService.js
    ↓
仅 WPS 版需要重新部署
复制新文件到 ~/.local/share/Kingsoft/wps/jsaddons/
```

---

## 文件依赖关系

```
GOWordAgent.WpsAddon/
    js/apiClient.js
        ↓ 调用 HTTP API
GOWordAgent.WpsService/
    Controllers/ProofreadController.cs
        ↓ 引用
GOWordAgent.Core/
    Services/ProofreadService.cs
        ↓ 引用
    Services/BaseLLMService.cs
        ↓ 调用 HTTP
DeepSeek/GLM/Ollama API
```

---

## 快速开始 Checklist

### 开发环境搭建

- [ ] 安装 .NET 8 SDK
- [ ] 克隆代码仓库
- [ ] 构建 Core 类库：`dotnet build GOWordAgent.Core`
- [ ] 构建服务：`dotnet build GOWordAgent.WpsService`

### 后端开发

- [ ] 修改 `GOWordAgent.Core/Services/` 下的服务类
- [ ] 修改 `GOWordAgent.WpsService/Controllers/` 下的控制器
- [ ] 运行测试：`dotnet test`（如果有测试项目）

### 前端开发

- [ ] 修改 `GOWordAgent.WpsAddon/js/` 下的 JS 文件
- [ ] 修改 `GOWordAgent.WpsAddon/index.html` 和 CSS
- [ ] 使用浏览器开发者工具调试 HTML/CSS

### 部署测试

- [ ] 执行 `dotnet publish` 生成 linux-x64 版本
- [ ] 复制到麒麟机器
- [ ] 执行 `./Scripts/install.sh`
- [ ] 按照 `TEST_GUIDE.md` 进行测试

---

*文档结束*
