# 智能校对 技术架构分析

## 一、项目概述

智能校对 是一个基于 **VSTO (Visual Studio Tools for Office)** 开发的 Word 插件，使用 **.NET Framework 4.8** 和 **WPF** 技术栈构建。

### 技术栈
| 层级 | 技术 |
|------|------|
| 前端 UI | WPF (XAML + C#) |
| Office 集成 | VSTO + Word Interop |
| 后端服务 | REST API (DeepSeek/GLM/Ollama) |
| 数据存储 | JSON + DPAPI 加密 |
| 并发控制 | async/await + SemaphoreSlim |

---

## 二、系统架构

### 2.1 整体架构图

```
┌─────────────────────────────────────────────────────────────┐
│                    Word Application                          │
│  ┌───────────────────────────────────────────────────────┐  │
│  │              CustomTaskPane (VSTO)                     │  │
│  │  ┌─────────────────────────────────────────────────┐  │  │
│  │  │         GOWordAgentPaneWpf (WPF UserControl)     │  │  │
│  │  │  ┌──────────────┐  ┌──────────────────────────┐ │  │  │
│  │  │  │   Chat View  │  │    Settings View         │ │  │  │
│  │  │  │              │  │  ┌─────────────────────┐ │ │  │  │
│  │  │  │  Message     │  │  │ AI Config Tab       │ │ │  │  │
│  │  │  │  List        │  │  └─────────────────────┘ │ │  │  │
│  │  │  │              │  │  ┌─────────────────────┐ │ │  │  │
│  │  │  │  Input Box   │  │  │ Proofread Config Tab│ │ │  │  │
│  │  │  └──────────────┘  │  └─────────────────────┘ │ │  │  │
│  │  └─────────────────────────────────────────────────┘  │  │
│  └───────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    Business Logic Layer                      │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │ProofreadSvc  │  │ConfigManager │  │WordDocumentHelper│  │
│  │- 并发控制    │  │- 加密存储    │  │- COM 封装        │  │
│  │- 缓存管理    │  │- 配置读写    │  │- 文档操作        │  │
│  └──────────────┘  └──────────────┘  └──────────────────┘  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    Service Layer                             │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌───────────┐  │
│  │DeepSeek  │  │   GLM    │  │  Ollama  │  │   Cache   │  │
│  │Service   │  │ Service  │  │ Service  │  │ Dictionary│  │
│  └──────────┘  └──────────┘  └──────────┘  └───────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### 2.2 核心类图

```
GOWordAgentPaneWpf (UserControl)
├── ILLMService _llmService
├── ProofreadService _proofreadService
├── List<object> _messageHistory
├── void StartProofread()
├── void ApplyProofreadToDocument()
└── void AddMessageBubble()

ILLMService (Interface)
├── Task<string> SendMessageAsync()
└── Task<string> SendProofreadMessageAsync()
    ▲
    │
    ├── BaseLLMService (Abstract)
    │   ├── HttpClient _httpClient
    │   └── abstract string BuildRequestBody()
    │       ▲
    │       ├── DeepSeekService
    │       ├── GLMService
    │       └── OllamaService

ProofreadService
├── ILLMService _llmService
├── SemaphoreSlim _semaphore (并发控制)
├── static Dictionary _globalCache (内存缓存)
├── Task<List<ParagraphResult>> ProofreadDocumentAsync()
└── string GenerateReport()

ConfigManager
├── static string _configPath
├── static AIConfig CurrentConfig
├── void SaveConfig() [DPAPI 加密]
└── void LoadConfig() [DPAPI 解密]
```

---

## 三、关键技术实现

### 3.1 并发分段处理

#### 分段策略
```csharp
// 1500 字/段，100 字重叠防止边界错误
const int TargetChunkSize = 1500;
const int OverlapSize = 100;

// 文本分段示例：
// [段落1: 0-1500] 
// [段落2: 1400-2900] (与段落1重叠100字)
// [段落3: 2800-4300] (与段落2重叠100字)
```

#### 并发控制
```csharp
// 使用信号量限制并发数（默认3）
private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(3);

// 每个段落处理时等待信号量
await _semaphore.WaitAsync(cancellationToken);
try {
    // 处理段落...
} finally {
    _semaphore.Release();
}
```

#### 优势
- **防止 API 限流**：控制并发请求数
- **提升效率**：多段并行处理
- **容错性**：单段失败不影响其他段

---

### 3.2 内存缓存机制

#### 缓存键设计
```csharp
// 基于内容的 SHA256 哈希作为缓存键
private string ComputeHash(string text) {
    using (var sha = SHA256.Create()) {
        var bytes = Encoding.UTF8.GetBytes(text);
        var hash = sha.ComputeHash(bytes);
        return Convert.ToBase64String(hash);
    }
}
```

#### 缓存策略
```csharp
// 静态缓存，跨实例共享
private static readonly Dictionary<string, ParagraphResult> _globalCache;

// 缓存命中检查
if (_globalCache.TryGetValue(cacheKey, out var cachedResult)) {
    return new ParagraphResult { IsCached = true, ... };
}
```

#### 特点
- **内存级缓存**：速度快，Word 关闭后清空
- **内容寻址**：相同内容直接返回缓存结果
- **并发安全**：使用 Dictionary（单线程访问设计）

---

### 3.3 Word 修订集成

#### 修订模式控制
```csharp
// 开启修订模式，记录修改历史
doc.TrackRevisions = true;
searchRange.Text = item.Modified;  // 产生删除线+下划线
doc.Comments.Add(searchRange, commentText);  // 添加批注

// 恢复原始状态（可选）
doc.TrackRevisions = oldTrackRevisions;
```

#### 修订效果
- **删除线**：标记原文
- **下划线**：标记建议文本
- **批注气泡**：显示类型、理由、原文、修改建议

---

### 3.4 数据安全

#### DPAPI 加密
```csharp
// 加密配置数据
byte[] encrypted = ProtectedData.Protect(
    Encoding.UTF8.GetBytes(json),
    null,
    DataProtectionScope.CurrentUser
);

// 解密配置数据
byte[] decrypted = ProtectedData.Unprotect(
    encrypted,
    null,
    DataProtectionScope.CurrentUser
);
```

#### 安全特性
- **用户级别加密**：仅当前 Windows 用户可解密
- **自动密钥管理**：Windows 自动处理加密密钥
- **无需手动输入密钥**：

---

## 四、设计模式

### 4.1 工厂模式
```csharp
// LLMServiceFactory 创建不同类型的服务
public static ILLMService CreateService(AIProvider provider, string apiKey, string apiUrl, string model) {
    switch (provider) {
        case AIProvider.DeepSeek: return new DeepSeekService(apiKey, apiUrl, model);
        case AIProvider.GLM: return new GLMService(apiKey, apiUrl, model);
        case AIProvider.Ollama: return new OllamaService(apiKey, apiUrl, model);
    }
}
```

### 4.2 策略模式
```csharp
// 不同的 AI 提供商实现相同的 ILLMService 接口
interface ILLMService {
    Task<string> SendMessageAsync(string message);
    Task<string> SendProofreadMessageAsync(string systemPrompt, string documentText);
}
```

### 4.3 观察者模式
```csharp
// ProofreadService 的进度回调
event EventHandler<ProofreadProgressArgs> OnProgress;

// UI 层订阅进度事件
proofreadService.OnProgress += (s, e) => {
    if (e.IsCached) AddMessageBubble($"第 {e.Result.Index + 1} 段", "📦 从缓存读取", false);
    else UpdateHeaderStatus(e.CurrentStatus, _primaryColor);
};
```

---

## 五、性能优化

### 5.1 异步处理
- 所有 I/O 操作使用 `async/await`
- UI 更新通过 `Dispatcher.InvokeAsync` 回到 UI 线程
- 防止 UI 卡顿

### 5.2 资源管理
```csharp
// 使用 try-finally 确保资源释放
await _semaphore.WaitAsync();
try {
    // 处理逻辑
} finally {
    _semaphore.Release();
}
```

### 5.3 COM 对象安全
```csharp
// 统一封装 COM 访问，添加空检查
try {
    var doc = Globals.ThisAddIn?.Application?.ActiveDocument;
    if (doc == null) return;
    // 操作文档...
} catch { /* 忽略 COM 错误 */ }
```

---

## 六、扩展性

### 6.1 添加新的 AI 提供商
1. 创建新类继承 `BaseLLMService`
2. 实现 `BuildRequestBody()` 方法
3. 在 `LLMServiceFactory` 中添加分支
4. 在 UI 中添加对应选项

### 6.2 自定义提示词
- 通过 `ConfigManager` 保存自定义提示词
- 支持两种模式（精准/全文）独立配置
- 提供默认提示词作为后备

### 6.3 缓存持久化
当前为内存缓存，如需持久化：
- 可将 `_globalCache` 改为文件存储（SQLite/JSON）
- 或使用分布式缓存（Redis）

---

## 七、总结

GOWordAgentAddIn 采用经典的分层架构，具有以下技术亮点：

1. **高并发**：Semaphore + 分段处理，支持大文档高效校对
2. **可扩展**：工厂模式 + 接口设计，易于添加新 AI 提供商
3. **安全性**：DPAPI 加密保护敏感配置
4. **用户体验**：WPF 流畅界面 + Word 原生修订无缝集成
5. **可维护性**：职责分离，代码结构清晰

项目代码量约 **2500 行**，适合作为 VSTO 插件开发的学习案例。
