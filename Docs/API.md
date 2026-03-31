# API 文档

## ILLMService 接口

所有 LLM 服务实现的统一接口。

```csharp
public interface ILLMService
{
    string ProviderName { get; }
    
    // 发送单条消息
    Task<string> SendMessageAsync(string userMessage, CancellationToken cancellationToken = default);
    
    // 发送带历史的消息
    Task<string> SendMessagesWithHistoryAsync(List<object> messages, CancellationToken cancellationToken = default);
    
    // 发送校对请求
    Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default);
}
```

## BaseLLMService

抽象基类，提供默认实现。

### 可覆盖成员

| 成员 | 说明 | 默认值 |
|------|------|--------|
| `ProofreadTimeoutSeconds` | 校对请求超时（秒） | 300 |
| `BuildRequestBodyDict` | 构建普通请求体 | OpenAI 格式 |
| `BuildProofreadRequestBodyDict` | 构建校对请求体 | temperature=0.1 |
| `ParseResponse` | 解析 API 响应 | choices[0].message.content |

### 使用示例

```csharp
// 创建服务
var service = new DeepSeekService(apiKey);

// 普通对话
string response = await service.SendMessageAsync("你好");

// 校对
string result = await service.SendProofreadMessageAsync(
    "你是校对专家...",  // system prompt
    "需要校对的文本..." // user content
);
```

## ProofreadService

校对核心服务。

### 构造函数

```csharp
public ProofreadService(
    ILLMService llmService,      // LLM 服务实例
    string systemPrompt,          // 系统提示词
    int concurrency = 5,          // 并发数（1-10）
    SegmenterConfig config = null // 分段配置
)
```

### 主要方法

```csharp
// 完整文档校对
Task<List<ParagraphResult>> ProofreadDocumentAsync(
    string documentText, 
    CancellationToken cancellationToken = default);

// 增量校对（只处理修改的段落）
Task<List<ParagraphResult>> ProofreadIncrementalAsync(
    string documentText,
    List<ParagraphResult> previousResults,
    CancellationToken cancellationToken = default);

// 生成报告
static string GenerateReport(
    List<ParagraphResult> results,
    int totalChars = 0,
    TimeSpan? elapsed = null,
    string providerName = null);

// 清除缓存
static void ClearCache();

// 获取缓存统计
static (int count, long totalBytes) GetCacheStats();
```

### 事件

```csharp
// 进度更新事件
public event EventHandler<ProofreadProgressArgs> OnProgress;

// 参数
public class ProofreadProgressArgs : EventArgs
{
    public int TotalParagraphs { get; set; }
    public int CompletedParagraphs { get; set; }
    public int CurrentIndex { get; set; }
    public string CurrentStatus { get; set; }
    public ParagraphResult Result { get; set; }
    public bool IsCompleted { get; set; }
    public int EstimatedRemainingSeconds { get; set; }
    public int CacheHitCount { get; set; }
}
```

### 优雅关闭

```csharp
// 使用 using 确保资源释放
using (var service = new ProofreadService(llm, prompt))
{
    var results = await service.ProofreadDocumentAsync(text);
}
// Dispose 时会：
// 1. 取消正在执行的任务
// 2. 等待任务完成（最多 5 秒）
// 3. 释放信号量
```

## WordDocumentService

Word COM 操作封装。

### 静态方法

```csharp
// 获取文档文本
static string GetDocumentText(Word.Application app);

// 高级版本（可控制是否包含表格、页眉页脚）
static string GetDocumentTextEx(
    Word.Application app,
    bool includeTables = true,
    bool includeHeadersFooters = false);
```

### 实例方法

```csharp
// 检查文档是否有效（COM 对象未释放）
bool IsDocumentValid();

// 查找文本位置（三级匹配策略）
(bool found, int start, int end) FindTextPosition(string text);

// 在范围内应用修订（带降级保护）
bool ApplyRevisionWithComment(
    string original, 
    string modified, 
    string commentText, 
    out int start, 
    out int end);

// 在指定位置应用修订（带降级保护）
bool ApplyRevisionAtRange(
    int start, 
    int end, 
    string original, 
    string modified, 
    string commentText,
    out int newStart, 
    out int newEnd);

// 导航到问题
void NavigateToIssue(ProofreadIssueItem item);

// 导航到指定范围（ScrollIntoView 降级保护）
bool NavigateToRange(int start, int end);

// 通过搜索导航（ScrollIntoView 降级保护）
bool NavigateBySearch(string originalText);
```

### Word 版本检测

```csharp
// 内部自动检测 Word 版本
private readonly int _wordVersion;

// 版本检测（如 16 = Word 2016/2019/365）
private static int GetWordVersion(Word.Application application)
{
    string versionString = application.Version; // "16.0.12345.67890"
    return int.Parse(versionString.Split('.')[0]);
}
```

## WordDocumentServiceFactory

文档服务工厂。

```csharp
// 从 ActiveDocument 创建（传统方式）
public static bool TryCreate(out WordDocumentService service, out string errorMessage);
public static WordDocumentService Create();

// 为特定文档创建（支持多文档场景）
public static bool TryCreateForDocument(
    Word.Application app, 
    Word.Document document, 
    out WordDocumentService service, 
    out string errorMessage);
```

## WordProofreadController

校对文档控制器（支持多文档）。

```csharp
public class WordProofreadController : IDisposable
{
    public WordProofreadController(Dispatcher dispatcher = null);
    
    // 获取当前文档文本
    public string GetDocumentText();
    
    // 应用校对结果到文档（多文档安全）
    public List<ProofreadIssueItem> ApplyProofreadToDocument(
        List<ProofreadIssueItem> items, 
        Action<string, string, bool, bool> addMessageCallback = null);
    
    // 在文档中定位到问题（多文档安全）
    public void NavigateToIssue(ProofreadIssueItem item);
    
    // 构建批注文本
    public static string BuildCommentText(ProofreadIssueItem item);
}
```

### 多文档支持

```csharp
// 内部自动处理文档切换
private bool TryGetDocumentService(out WordDocumentService service, out string errorMessage)
{
    var activeDoc = app.ActiveDocument;
    
    // 如果切换到不同文档，自动重新绑定
    if (_documentService == null || !IsSameDocument(_boundDocument, activeDoc))
    {
        _documentService?.Dispose();
        WordDocumentServiceFactory.TryCreateForDocument(app, activeDoc, out _documentService, ...);
        _boundDocument = activeDoc;
    }
}
```

## ConfigManager

配置管理（静态类）。

```csharp
// 当前配置
public static AIConfig CurrentConfig { get; }

// 保存配置（DPAPI 加密 + HMAC 完整性验证）
public static void SaveConfig(AIConfig config);

// 加载配置（验证 HMAC 完整性）
public static void LoadConfig();

// 获取指定提供商配置
public static ProviderConfig GetProviderConfig(AIProvider provider);

// 获取校验配置
public static (string mode, string prompt) GetProofreadConfig();

// 获取隐私同意日期
public static string PrivacyConsentLastShownDate { get; set; }
```

### 配置安全

```csharp
// 加密：DPAPI（用户级加密）
// 完整性：HMAC-SHA256（防篡改）
private static readonly string HmacKeyFile = Path.Combine(ConfigDir, "config.key");

// 存储格式：HMAC(32字节) + EncryptedData
```

## HttpClientFactory

HttpClient 工厂（静态类）。

```csharp
// 创建带认证的客户端
HttpClient CreateAuthenticatedClient(
    string apiKey, 
    string apiUrl, 
    int timeoutSeconds = 120);

// 清理缓存（DNS 变更时调用）
void ClearCache();

// 获取缓存数量（调试用）
int CachedClientCount { get; }
```

### 共享连接池

```csharp
// 静态共享 Handler
private static readonly HttpClientHandler _sharedHandler = new HttpClientHandler
{
    MaxConnectionsPerServer = 20,
    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
};

// 每个服务独立 Client 实例，但共享 Handler 连接池
return new HttpClient(_sharedHandler, disposeHandler: false);
```

## ProofreadCacheManager

缓存管理（静态类）。

```csharp
// 尝试获取缓存结果
public static bool TryGetCachedResult(string text, int index, out ParagraphResult result);

// 存储结果
public static void StoreResult(string text, ParagraphResult result);

// 计算内容哈希（SHA256 优化）
public static string ComputeHash(string text);

// 清除缓存
public static void ClearCache();

// 获取统计
public static (int Count, long EstimatedBytes) GetCacheStats();
```

### SHA256 优化

```csharp
// 静态实例 + 锁，避免高并发时频繁创建
private static readonly SHA256 _sha256 = SHA256.Create();
private static readonly object _shaLock = new object();

public static string ComputeHash(string text)
{
    var bytes = Encoding.UTF8.GetBytes(text);
    byte[] hash;
    lock (_shaLock)  // 细粒度锁
    {
        hash = _sha256.ComputeHash(bytes);
    }
    return Convert.ToBase64String(hash);
}
```

## 异常处理

### LLMServiceException

```csharp
try
{
    var result = await service.SendProofreadMessageAsync(...);
}
catch (LLMServiceException ex)
{
    // 获取友好的错误信息
    string message = ex.GetFriendlyErrorMessage();
    // 可能值：
    // - [DeepSeek] API Key 无效或已过期
    // - [GLM] 请求过于频繁，请稍后重试
    // - [Ollama] 连接失败，请确认服务是否运行
}
```

### COMException 处理

```csharp
// Word API 调用失败时自动降级
try
{
    _document.TrackRevisions = true;
}
catch (COMException)
{
    // 降级为普通替换（无修订标记）
    range.Text = modified;
}
```

## 数据模型

### ParagraphResult

```csharp
public class ParagraphResult
{
    public int Index { get; set; }              // 段落索引
    public string OriginalText { get; set; }    // 原始文本
    public string ResultText { get; set; }      // AI 返回的完整结果
    public bool IsCompleted { get; set; }       // 是否完成
    public bool IsCached { get; set; }          // 是否来自缓存
    public DateTime ProcessTime { get; set; }   // 处理时间
    public long ElapsedMs { get; set; }         // 耗时（毫秒）
    public List<ProofreadIssueItem> Items { get; set; } // 解析出的问题
}
```

### ProofreadIssueItem

```csharp
public class ProofreadIssueItem
{
    public int Index { get; set; }          // 问题序号
    public string Type { get; set; }        // 问题类型（错别字、语病等）
    public string Original { get; set; }    // 原文
    public string Modified { get; set; }    // 修改建议
    public string Reason { get; set; }      // 修改理由
    public string Severity { get; set; }    // 严重程度（high/medium/low）
    public int DocumentStart { get; set; }  // 文档起始位置
    public int DocumentEnd { get; set; }    // 文档结束位置
}
```
