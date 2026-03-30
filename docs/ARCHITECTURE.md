# 架构说明

## 项目结构

```
GOWordAgentAddIn/
├── Models/                  # 数据模型
│   └── ProofreadModels.cs   # ParagraphResult, ProofreadIssueItem 等
├── ViewModels/              # MVVM 视图模型
│   ├── ChatMessageViewModel.cs
│   ├── ComplexMessageViewModel.cs
│   ├── Converters.cs
│   └── RelayCommand.cs
├── Services/                # 业务服务
│   ├── ProofreadService.cs          # 校对核心服务
│   ├── WordDocumentService.cs       # Word COM 操作
│   ├── WordProofreadController.cs   # 校对文档交互（多文档支持）
│   ├── BaseLLMService.cs            # LLM 服务基类
│   ├── DeepSeekService.cs
│   ├── GLMService.cs
│   ├── OllamaService.cs
│   └── LLMServiceFactory.cs
├── Http/                    # HTTP 相关
│   └── HttpClientFactory.cs # HttpClient 共享工厂
├── Utils/                   # 工具类
│   ├── ConfigManager.cs              # 配置管理（DPAPI + HMAC）
│   ├── ProofreadCacheManager.cs      # 缓存管理（SHA256 优化）
│   ├── ProofreadIssueParser.cs
│   ├── ProofreadResultRenderer.cs    # 静态 Brush 复用
│   ├── DocumentSegmenter.cs
│   └── LLMRequestLogger.cs
└── UI/                      # 用户界面
    ├── GOWordAgentPaneWpf.xaml
    └── GOWordAgentPaneWpf.xaml.cs
```

## 核心架构

### MVVM 模式

```
View (XAML) <-> ViewModel <-> Model
     ↓
DataTemplate 自动渲染
```

- **View**: GOWordAgentPaneWpf.xaml + DataTemplates
- **ViewModel**: ChatMessageViewModel, ComplexMessageViewModel
- **Model**: ParagraphResult, ProofreadIssueItem

### LLM 服务继承体系

```
ILLMService (接口)
    └─ BaseLLMService (抽象基类)
         ├─ DeepSeekService  → 使用默认实现
         ├─ GLMService       → 覆盖 BuildProofreadRequestBodyDict (enable_thinking)
         └─ OllamaService    → 覆盖 BuildRequestBody/BuildProofreadRequestBody, 错误处理
```

### 校对流程

```
1. 获取文档文本
   WordProofreadController.GetDocumentText()
   
2. 分段处理
   DocumentSegmenter.SplitIntoParagraphs()
   
3. 并行校对
   ProofreadService.ProofreadDocumentAsync()
   ├── 检查缓存 (ProofreadCacheManager - SHA256 静态优化)
   ├── 调用 LLM (BaseLLMService.SendProofreadMessageAsync)
   │   ├── 构建请求体 (BuildProofreadRequestBodyDict)
   │   ├── HTTP 请求 (HttpClientFactory)
   │   └── 解析响应 (ParseResponse)
   ├── 解析问题 (ProofreadIssueParser.ParseProofreadItems)
   └── 存储缓存
   
4. 生成报告
   ProofreadService.GenerateReport()
   
5. 应用到文档
   WordProofreadController.ApplyProofreadToDocument()
   └── WordDocumentService.ApplyRevisionAtRange()
```

## 关键技术点

### 1. HttpClient 共享

```csharp
// HttpClientFactory 使用共享 Handler
private static readonly HttpClientHandler _sharedHandler = new HttpClientHandler
{
    MaxConnectionsPerServer = 20,
    // ... SSL, 压缩配置
};

// 每个服务实例创建独立的 HttpClient，但共享连接池
public static HttpClient CreateAuthenticatedClient(string apiKey, string apiUrl, int timeoutSeconds = 120)
{
    return new HttpClient(_sharedHandler, disposeHandler: false)
    {
        Timeout = TimeSpan.FromSeconds(timeoutSeconds)
    };
}
```

### 2. 并发控制（优雅关闭）

```csharp
// ProofreadService 使用 Semaphore 控制并发
private readonly SemaphoreSlim _semaphore;
private readonly CancellationTokenSource _disposeCts = new CancellationTokenSource();
private long _activeTaskCount = 0;

// 默认并发数 5，可在构造函数调整
public ProofreadService(ILLMService llmService, string systemPrompt, int concurrency = 5)
{
    _semaphore = new SemaphoreSlim(concurrency);
}

// Dispose 时优雅关闭
protected virtual void Dispose(bool disposing)
{
    _disposeCts?.Cancel();  // 1. 通知任务退出
    // 2. 等待任务完成（最多 5 秒）
    while (_activeTaskCount > 0) Thread.Sleep(50);
    _semaphore?.Dispose();   // 3. 释放信号量
}
```

### 3. 缓存机制（SHA256 优化）

```csharp
// 静态 SHA256 实例，避免高并发时频繁创建
private static readonly SHA256 _sha256 = SHA256.Create();
private static readonly object _shaLock = new object();

public static string ComputeHash(string text)
{
    var bytes = Encoding.UTF8.GetBytes(text);
    byte[] hash;
    lock (_shaLock)  // 细粒度锁，减少竞争
    {
        hash = _sha256.ComputeHash(bytes);
    }
    return Convert.ToBase64String(hash);
}
```

### 4. COM 对象生命周期管理

```csharp
// 统一使用 try-finally + ReleaseComObject
Word.Range range = null;
Word.Find find = null;
try
{
    range = _document.Range(start, end);
    find = range.Find;
    // ... 操作
}
finally
{
    if (find != null) Marshal.ReleaseComObject(find);
    if (range != null) Marshal.ReleaseComObject(range);
}
```

**重要规则**:
- `app.ActiveDocument` 是外部引用，**不释放**
- 自创建的 `Range`、`Find`、`Comment` 必须释放
- 使用 `try-finally` 确保异常时也能释放

### 5. Word 版本兼容性

```csharp
// 自动检测 Word 版本
private static int GetWordVersion(Word.Application application)
{
    string versionString = application.Version; // "16.0.12345.67890"
    var parts = versionString.Split('.');
    return int.Parse(parts[0]); // 16 = Word 2016/2019/365
}

// API 降级策略
try
{
    _document.TrackRevisions = true;
    range.Text = modified;
}
catch (COMException)
{
    // 降级：不使用修订模式
    range.Text = modified;
}
```

**兼容性矩阵**:

| API | Word 2016+ | Word 2013 | Word 2010 | 降级策略 |
|-----|:----------:|:---------:|:---------:|----------|
| TrackRevisions | ✅ | ✅ | ✅ | 普通替换 |
| Comments.Add | ✅ | ✅ | ✅ | 跳过批注 |
| ScrollIntoView | ✅ | ✅ | ⚠️ | try-catch |

### 6. 多文档支持

```csharp
public class WordProofreadController : IDisposable
{
    private WordDocumentService _documentService;
    private Word.Document _boundDocument;
    
    private bool TryGetDocumentService(out WordDocumentService service, out string errorMessage)
    {
        var activeDoc = app.ActiveDocument;
        
        // 如果切换到不同文档，重新绑定
        if (_documentService == null || !IsSameDocument(_boundDocument, activeDoc))
        {
            _documentService?.Dispose();
            WordDocumentServiceFactory.TryCreateForDocument(app, activeDoc, out _documentService, ...);
            _boundDocument = activeDoc;
        }
        
        // 验证文档是否仍有效
        if (!_documentService.IsDocumentValid()) { ... }
    }
    
    private bool IsSameDocument(Word.Document doc1, Word.Document doc2)
    {
        return doc1.FullName == doc2.FullName;
    }
}
```

### 7. DPI 感知（高分辨率屏幕优化）

```xml
<!-- app.manifest -->
<application xmlns="urn:schemas-microsoft-com:asm.v3">
  <windowsSettings>
    <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>
    <dpiAwareness xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">system</dpiAwareness>
  </windowsSettings>
</application>
```

```csharp
// ThisAddIn.cs
TextOptions.TextFormattingModeProperty.OverrideMetadata(
    typeof(Window),
    new FrameworkPropertyMetadata(TextFormattingMode.Display));
```

### 8. 错误处理

```csharp
// 自定义异常，包含详细错误信息
public class LLMServiceException : Exception
{
    public HttpStatusCode? StatusCode { get; }
    public string ProviderName { get; }
    
    public string GetFriendlyErrorMessage()
    {
        // 根据状态码返回友好的错误信息
    }
}
```

## 数据流

```
Word 文档
    ↓
WordProofreadController (多文档安全)
    ↓
WordDocumentService.GetDocumentText()
    ↓
DocumentSegmenter.SplitIntoParagraphs()
    ↓
┌─────────────────────────────────────────┐
│  段落 1  │  段落 2  │  段落 3  │  ...   │  ← 并行处理（Semaphore 控制）
└─────────────────────────────────────────┘
    ↓
ProofreadCacheManager.ComputeHash() (SHA256 静态优化)
    ↓
BaseLLMService.SendProofreadMessageAsync()
    ↓
AI 提供商 API
    ↓
ProofreadIssueParser.ParseProofreadItems()
    ↓
ParagraphResult[]
    ↓
ProofreadService.GenerateReport()
    ↓
WordProofreadController.ApplyProofreadToDocument()
    ↓
Word 文档（修订模式 / 降级为普通替换）
```

## 性能优化要点

1. **Brush 复用**: `SolidColorBrush` 使用静态只读字段，调用 `Freeze()`
2. **SHA256 复用**: 静态实例 + 锁，避免频繁创建哈希算法实例
3. **HttpClient 共享**: 共享 Handler，独立 Client 实例
4. **缓存机制**: LRU 淘汰，基于内容哈希避免重复计算
5. **COM 释放**: 及时释放 COM 对象，减少 Word 进程内存占用
