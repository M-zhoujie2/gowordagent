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
// 检查文档是否有效
bool IsDocumentValid();

// 查找文本位置
(bool found, int start, int end) FindTextPosition(string text);

// 在范围内应用修订
bool ApplyRevisionWithComment(
    string original, 
    string modified, 
    string commentText, 
    out int start, 
    out int end);

// 在指定位置应用修订
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
```

## ConfigManager

配置管理（静态类）。

```csharp
// 当前配置
public static AIConfig CurrentConfig { get; }

// 保存配置
public static void SaveConfig(AIConfig config);

// 加载配置
public static void LoadConfig();

// 获取指定提供商配置
public static ProviderConfig GetProviderConfig(AIProvider provider);

// 获取校验配置
public static (string mode, string prompt) GetProofreadConfig();
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
