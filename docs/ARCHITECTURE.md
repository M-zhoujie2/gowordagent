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
│   ├── WordProofreadController.cs   # 校对文档交互
│   ├── BaseLLMService.cs            # LLM 服务基类
│   ├── DeepSeekService.cs
│   ├── GLMService.cs
│   ├── OllamaService.cs
│   └── LLMServiceFactory.cs
├── Http/                    # HTTP 相关
│   └── HttpClientFactory.cs # HttpClient 共享工厂
├── Utils/                   # 工具类
│   ├── ConfigManager.cs
│   ├── ProofreadCacheManager.cs
│   ├── ProofreadIssueParser.cs
│   ├── ProofreadResultRenderer.cs
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
   ├── 检查缓存 (ProofreadCacheManager)
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

### HttpClient 共享

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

### 并发控制

```csharp
// ProofreadService 使用 Semaphore 控制并发
private readonly SemaphoreSlim _semaphore;

// 默认并发数 5，可在构造函数调整
public ProofreadService(ILLMService llmService, string systemPrompt, int concurrency = 5)
{
    _semaphore = new SemaphoreSlim(concurrency);
}
```

### 缓存机制

```csharp
// 静态缓存，跨校对会话有效
private static readonly Dictionary<string, ParagraphResult> _globalCache = new Dictionary<string, ParagraphResult>();

// LRU 淘汰策略，最大 1000 条
private const int MaxCacheSize = 1000;
```

### 错误处理

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
WordDocumentService.GetDocumentText()
    ↓
DocumentSegmenter.SplitIntoParagraphs()
    ↓
┌─────────────────────────────────────────┐
│  段落 1  │  段落 2  │  段落 3  │  ...   │  ← 并行处理（Semaphore 控制）
└─────────────────────────────────────────┘
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
Word 文档（修订模式）
```
