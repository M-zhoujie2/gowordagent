# GOWordAgent 银河麒麟 V10 改造完成报告

> **日期**：2026-04-13  
> **状态**：✅ 完成

---

## 改造完成情况

### 1. GOWordAgent.Core (.NET 8 类库) ✅

| 文件 | 说明 | 状态 |
|------|------|------|
| `GOWordAgent.Core.csproj` | 项目文件 | ✅ |
| `Models/ProofreadModels.cs` | 数据模型 | ✅ |
| `Services/ILLMService.cs` | 服务接口 | ✅ |
| `Services/BaseLLMService.cs` | LLM 基类 | ✅ |
| `Services/DeepSeekService.cs` | DeepSeek 适配 | ✅ |
| `Services/GLMService.cs` | GLM 适配 | ✅ |
| `Services/OllamaService.cs` | Ollama 适配 | ✅ |
| `Services/LLMServiceFactory.cs` | 服务工厂 | ✅ |
| `Services/LLMServiceException.cs` | 异常定义 | ✅ |
| `Services/ProofreadService.cs` | 校对服务核心 | ✅ |
| `Services/DocumentSegmenter.cs` | 文档分段器 | ✅ |
| `Services/ProofreadCacheManager.cs` | 缓存管理 | ✅ |
| `Services/ProofreadIssueParser.cs` | 结果解析 | ✅ |
| `Config/ConfigManager.cs` | 跨平台配置管理 | ✅ |

**代码统计**：
- 核心服务类：12 个文件
- 总代码量：约 8,500 行（含注释）
- 修改点：移除 WPF Dispatcher 依赖，改为事件驱动

### 2. GOWordAgent.WpsService (.NET 8 Minimal API) ✅

| 文件 | 说明 | 状态 |
|------|------|------|
| `GOWordAgent.WpsService.csproj` | 项目文件（Self-Contained） | ✅ |
| `Program.cs` | 服务入口 | ✅ |
| `Controllers/ProofreadController.cs` | API 控制器 | ✅ |

**API 端点**：
- `POST /api/proofread` - 执行校对
- `GET /api/proofread/health` - 健康检查
- `GET /api/proofread/config` - 获取配置
- `POST /api/proofread/config` - 保存配置

### 3. GOWordAgent.WpsAddon (WPS 加载项) ✅

| 文件 | 说明 | 状态 |
|------|------|------|
| `package.json` | wpsjs 配置 | ✅ |
| `index.html` | 主页面 | ✅ |
| `main.js` | 入口脚本 | ✅ |
| `css/style.css` | 样式表 | ✅ |
| `js/apiClient.js` | 后端通信 | ✅ |
| `js/documentService.js` | WPS JS API 封装 | ✅ |
| `js/proofreadService.js` | 校对工作流 | ✅ |
| `js/uiController.js` | UI 控制 | ✅ |

**功能特性**：
- AI 配置面板（提供商、API Key、模型、校验模式）
- 校对结果列表（严重程度标记、点击定位）
- 实时进度显示
- 连接状态指示

### 4. 部署脚本 ✅

| 文件 | 说明 | 状态 |
|------|------|------|
| `Scripts/install.sh` | 安装脚本 | ✅ |
| `Scripts/uninstall.sh` | 卸载脚本 | ✅ |
| `Scripts/gowordagent.service` | systemd 服务文件 | ✅ |

### 5. 文档 ✅

| 文件 | 说明 | 状态 |
|------|------|------|
| `Docs/KYLIN_V10_MIGRATION_RECORD.md` | 详细改造记录 | ✅ |
| `Docs/CODE_OPTIMIZATION_PLAN.md` | 代码优化方案 | ✅ |
| `KYLIN_V10_BUILD.md` | 构建指南 | ✅ |
| `MIGRATION_COMPLETE.md` | 本报告 | ✅ |

---

## 关键改造点

### 1. 跨平台配置加密

**Windows（原有）**：
```csharp
// DPAPI 加密
byte[] encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
```

**Linux（新增）**：
```csharp
// AES-GCM + /etc/machine-id
using var aes = new AesGcm(key, 16);
aes.Encrypt(nonce, plainData, cipherData, tag);
```

### 2. 移除 WPF 依赖

**原代码**：
```csharp
private readonly Dispatcher _dispatcher;
await _dispatcher.InvokeAsync(() => OnProgress?.Invoke(this, args));
```

**新代码**：
```csharp
public event EventHandler<ProofreadProgressArgs>? OnProgress;
OnProgress?.Invoke(this, args); // 直接触发事件
```

### 3. 字符偏移量定位

**方案**：前后端统一使用字符偏移量定位，替代 Word Range.Find

**后端返回**：
```json
{
  "startOffset": 1250,
  "endOffset": 1256,
  "original": "错误的词",
  "suggestion": "正确的词"
}
```

**前端应用**：
```javascript
var range = doc.Range(startOffset, endOffset);
range.Delete();
range.InsertAfter(replacement);
```

### 4. Self-Contained 部署

```xml
<SelfContained>true</SelfContained>
<RuntimeIdentifier>linux-x64</RuntimeIdentifier>
<PublishSingleFile>true</PublishSingleFile>
```

**优势**：
- 零运行时依赖
- 单文件可执行
- 60-100MB 包体积（信创桌面可接受）

---

## 项目结构总览

```
GOWordAgent/
├── GOWordAgentAddIn/              # 【保留】原有 Word VSTO
│   └── ... (保持不变)
│
├── GOWordAgent.Core/              # 【新增】.NET 8 类库
│   ├── Config/
│   │   └── ConfigManager.cs       # 跨平台配置
│   ├── Models/
│   │   └── ProofreadModels.cs     # 数据模型
│   └── Services/
│       ├── BaseLLMService.cs      # LLM 基类
│       ├── ProofreadService.cs    # 校对服务
│       ├── DocumentSegmenter.cs   # 文档分段
│       ├── ProofreadCacheManager.cs # 缓存
│       ├── ProofreadIssueParser.cs  # 解析器
│       └── ...                    # 其他服务
│
├── GOWordAgent.WpsService/        # 【新增】后端服务
│   ├── Controllers/
│   │   └── ProofreadController.cs # API 控制器
│   └── Program.cs                 # 服务入口
│
├── GOWordAgent.WpsAddon/          # 【新增】WPS 加载项
│   ├── css/style.css              # 样式
│   ├── js/                        # 脚本
│   │   ├── apiClient.js           # 后端通信
│   │   ├── documentService.js     # 文档操作
│   │   ├── proofreadService.js    # 工作流
│   │   └── uiController.js        # UI 控制
│   ├── index.html                 # 主页面
│   └── package.json               # 加载项配置
│
└── Scripts/                       # 【新增】部署脚本
    ├── install.sh
    ├── uninstall.sh
    └── gowordagent.service
```

---

## 后续步骤

### 1. 构建测试

```bash
# Windows 开发环境
cd GOWordAgent.WpsService
dotnet publish -c Release -r linux-x64 --self-contained true
```

### 2. Day 1 PoC 验证

在银河麒麟 V10 实机验证：
- [ ] WPS 加载项能否访问 `http://127.0.0.1:PORT`
- [ ] `fetch` 或 `XMLHttpRequest` 是否被拦截
- [ ] 后端服务能否正常启动

### 3. 完整流程测试

- [ ] 配置保存/加载
- [ ] 文档文本提取
- [ ] 校对执行
- [ ] 结果写回（修订+批注）
- [ ] 点击定位

---

## 技术决策记录

| 决策 | 选择 | 理由 |
|------|------|------|
| 前端框架 | 原生 HTML/JS | 避免构建链复杂性，兼容性好 |
| JSON 库 | Newtonsoft.Json | 现有代码稳定，不迁移到 System.Text.Json |
| 部署方式 | Self-Contained | 零依赖，适合信创环境 |
| 文本定位 | 字符偏移量 | 避免 Find.Execute 重复匹配问题 |
| 加密方案 | AES-GCM + machine-id | Linux 原生兼容，无需额外库 |
| 进程管理 | systemd 用户服务 | 稳定可靠，用户级权限 |

---

## 风险与应对

| 风险 | 概率 | 应对 |
|------|------|------|
| WPS WebView 阻止 localhost 访问 | 中 | Day 1 PoC 验证，失败则切换 LibreOffice 方案 |
| TrackRevisions 行为不一致 | 中 | 实现降级逻辑（直接替换+批注） |
| 大批量注导致卡顿 | 中 | 限制单次写入数量（最多 50 条） |
| .NET 8 在麒麟运行异常 | 极低 | x86_64 是一级支持平台 |

---

**改造完成时间**：2026-04-13  
**预计联调时间**：Day 1-10

*文档结束*
