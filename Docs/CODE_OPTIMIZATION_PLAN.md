# GOWordAgentAddIn 代码优化方案

> 生成日期：2026-04-13  
> 版本：v1.0  
> 状态：待评审

---

## 目录

1. [概述](#概述)
2. [架构优化](#架构优化)
3. [性能优化](#性能优化)
4. [文本校验专项优化](#文本校验专项优化)
5. [可靠性优化](#可靠性优化)
6. [代码质量优化](#代码质量优化)
7. [安全优化](#安全优化)
8. [实施路线图](#实施路线图)

---

## 概述

### 项目现状

GOWordAgentAddIn 是一个基于 VSTO 的 Word 插件，提供 AI 智能校对功能。项目采用分层架构，支持多种 LLM 提供商（DeepSeek、GLM、Ollama），具备并发处理、缓存机制、MVVM 架构等特性。

### 优化目标

- **性能提升**：减少响应时间，降低内存占用
- **可靠性增强**：提高容错能力，减少崩溃风险
- **可维护性**：代码结构清晰，便于后续扩展
- **用户体验**：更快的校对速度，更准确的定位

### 优先级说明

| 优先级 | 说明 | 建议处理时间 |
|--------|------|-------------|
| 🔴 高 | 影响稳定性或性能的关键问题 | 立即处理 |
| 🟡 中 | 影响体验或长期维护的问题 | 1-2 个迭代 |
| 🟢 低 | 代码质量改进或增强功能 | 按需处理 |

---

## 架构优化

### 1. 代码文件拆分（🔴 高）

#### 问题
`GOWordAgentPaneWpf.xaml.cs` 超过 1200 行，职责混杂，难以维护。

#### 优化方案
按功能拆分为多个 Partial Class 或独立 Service：

```
GOWordAgentPaneWpf.xaml.cs          -- 仅保留 UI 事件绑定
GOWordAgentPaneWpf.Chat.cs          -- 聊天功能
GOWordAgentPaneWpf.Proofread.cs     -- 校对功能
GOWordAgentPaneWpf.Config.cs        -- 配置管理
GOWordAgentPaneWpf.Navigation.cs    -- 导航功能
```

#### 实施步骤
1. 创建 Partial Class 文件
2. 按功能区域移动方法
3. 提取公共逻辑到基类或 Helper

---

### 2. 消除代码重复（🔴 高）

#### 问题
`CreateIssueButton` 方法在以下两个文件中重复实现：
- `GOWordAgentPaneWpf.xaml.cs` (第 1072 行)
- `ProofreadResultRenderer.cs` (第 211 行)

#### 优化方案
统一使用 `ProofreadResultRenderer`，删除 `GOWordAgentPaneWpf` 中的重复实现。

```csharp
// 修改 GOWordAgentPaneWpf.xaml.cs
private void AddProofreadResultBubble(string reportTitle, string reportContent, 
    List<ProofreadIssueItem> items, List<ParagraphResult> paragraphResults)
{
    var renderer = new ProofreadResultRenderer(
        MessagesPanel, 
        ChatScrollViewer,
        _aiBubbleColor,
        _textPrimaryColor,
        _textSecondaryColor,
        item => _wordController.NavigateToIssue(item));
    
    renderer.AddProofreadResultBubble(reportTitle, reportContent, items, paragraphResults);
}
```

---

### 3. 服务接口抽象（🟡 中）

#### 问题
LLM 服务扩展需要修改工厂类，不符合开闭原则。

#### 优化方案
引入依赖注入或服务发现机制：

```csharp
public interface ILLMServiceProvider
{
    string ProviderName { get; }
    bool CanHandle(AIProvider provider);
    ILLMService CreateService(string apiKey, string apiUrl, string model);
}

public class LLMServiceRegistry
{
    private static readonly List<ILLMServiceProvider> _providers = new List<ILLMServiceProvider>();
    
    public static void Register(ILLMServiceProvider provider) => _providers.Add(provider);
    
    public static ILLMService CreateService(AIProvider provider, string apiKey, string apiUrl, string model)
    {
        var serviceProvider = _providers.FirstOrDefault(p => p.CanHandle(provider));
        return serviceProvider?.CreateService(apiKey, apiUrl, model);
    }
}
```

---

## 性能优化

### 1. SHA256 线程安全修复（🔴 高）

#### 问题
`ProofreadCacheManager` 中使用静态 `SHA256` 实例加锁，高并发时可能成为瓶颈且存在线程安全问题。

**当前代码：**
```csharp
private static readonly SHA256 _sha256 = SHA256.Create();
lock (_shaLock) { hash = _sha256.ComputeHash(bytes); }
```

#### 优化方案
使用实例方法或 `SHA256.Create()` 每次创建新实例（.NET Framework 4.8 推荐）：

```csharp
public static string ComputeHash(string text, string mode = null)
{
    if (string.IsNullOrEmpty(text)) return string.Empty;
    
    var content = string.IsNullOrEmpty(mode) ? text : $"[{mode}]{text}";
    var bytes = Encoding.UTF8.GetBytes(content);
    
    // 每次创建新实例，避免线程安全问题
    using (var sha256 = SHA256.Create())
    {
        var hash = sha256.ComputeHash(bytes);
        return Convert.ToBase64String(hash);
    }
}
```

#### 预期收益
- 消除线程竞争
- 简化代码逻辑
- 微秒级性能差异可忽略

---

### 2. HttpClient DNS 刷新问题（🔴 高）

#### 问题
`HttpClientFactory` 使用静态共享的 `HttpClientHandler`，DNS 变更时不会自动刷新。

#### 优化方案
实现带过期时间的 Handler 刷新机制：

```csharp
public static class HttpClientFactory
{
    private static readonly ConcurrentDictionary<string, (HttpClient Client, DateTime Created)> _clients = 
        new ConcurrentDictionary<string, (HttpClient, DateTime)>();
    
    private static readonly TimeSpan _clientLifetime = TimeSpan.FromMinutes(30);
    
    public static HttpClient GetClient(string baseAddress)
    {
        var normalizedKey = NormalizeBaseAddress(baseAddress);
        
        // 检查是否需要刷新
        if (_clients.TryGetValue(normalizedKey, out var existing))
        {
            if (DateTime.Now - existing.Created < _clientLifetime)
            {
                return existing.Client;
            }
            
            // 过期，创建新的
            existing.Client.Dispose();
        }
        
        var newClient = CreateNewClient(normalizedKey);
        _clients[normalizedKey] = (newClient, DateTime.Now);
        return newClient;
    }
}
```

---

### 3. 缓存内存上限（🔴 高）

#### 问题
`ProofreadCacheManager` 使用静态字典，Word 插件生命周期内持续增长，可能导致内存泄漏。

#### 优化方案
添加内存上限和定期清理：

```csharp
public static class ProofreadCacheManager
{
    private const long MaxCacheBytes = 50 * 1024 * 1024; // 50MB
    private static long _currentCacheBytes = 0;
    
    public static void StoreResult(string text, ParagraphResult result, string mode = null)
    {
        var estimatedBytes = (text?.Length ?? 0) + (result?.ResultText?.Length ?? 0) * 2;
        
        // 检查是否超出内存限制
        if (Interlocked.Add(ref _currentCacheBytes, estimatedBytes) > MaxCacheBytes)
        {
            EvictOldestEntries(MaxCacheSize / 4); // 清理 25%
        }
        // ... 原有逻辑
    }
}
```

---

### 4. UI 更新频率优化（🟡 中）

#### 问题
`ProofreadService` 中每完成一个段落就更新 UI，大量段落时 UI 刷新过于频繁。

#### 优化方案
引入节流机制：

```csharp
public class ProofreadService
{
    private DateTime _lastProgressUpdate = DateTime.MinValue;
    private readonly TimeSpan _progressThrottleInterval = TimeSpan.FromMilliseconds(200);
    
    private void ReportProgressThrottled(ProofreadProgressArgs args)
    {
        var now = DateTime.Now;
        if (now - _lastProgressUpdate < _progressThrottleInterval && !args.IsCompleted)
        {
            return; // 跳过过于频繁的更新
        }
        
        _lastProgressUpdate = now;
        ReportProgress(args);
    }
}
```

---

### 5. 流式响应支持（🟢 低）

#### 问题
当前实现等待 LLM 完整响应后才更新 UI，用户感知延迟较长。

#### 优化方案
添加流式响应支持（需要 LLM 提供商支持）：

```csharp
public async Task StreamProofreadMessageAsync(
    string systemContent, 
    string userContent, 
    Action<string> onChunkReceived,
    CancellationToken cancellationToken = default)
{
    var request = BuildStreamRequest(systemContent, userContent);
    
    using (var response = await _httpClient.SendAsync(
        request, 
        HttpCompletionOption.ResponseHeadersRead, 
        cancellationToken))
    {
        using (var stream = await response.Content.ReadAsStreamAsync())
        using (var reader = new StreamReader(stream))
        {
            while (!reader.EndOfStream && !cancellationToken.IsCancellationRequested)
            {
                var line = await reader.ReadLineAsync();
                var chunk = ParseStreamChunk(line);
                onChunkReceived?.Invoke(chunk);
            }
        }
    }
}
```

---

## 文本校验专项优化

### 1. 智能语义分段（🔴 高）

#### 问题
当前分段仅按句子切分，未考虑语义完整性，可能导致上下文丢失。

#### 优化方案
实现语义感知分段：

```csharp
public class SmartDocumentSegmenter
{
    public List<Segment> SplitIntoSemanticChunks(string text)
    {
        var segments = new List<Segment>();
        
        // 先按段落分割（保留换行语义）
        var paragraphs = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        
        foreach (var para in paragraphs)
        {
            if (string.IsNullOrWhiteSpace(para)) continue;
            
            if (para.Length <= _config.TargetChunkSize)
            {
                segments.Add(CreateSegment(para, SegmentType.Paragraph));
                continue;
            }
            
            // 长段落按语义边界分割
            segments.AddRange(SplitParagraphSmart(para));
        }
        
        return segments;
    }
    
    private List<Segment> SplitParagraphSmart(string paragraph)
    {
        var segments = new List<Segment>();
        // 按语义边界分割：句号 > 分号/逗号 > 固定长度
        var boundaries = new[] { "。", "；", "，" };
        // 实现分割逻辑...
        return segments;
    }
}

public class Segment
{
    public string Text { get; set; }
    public SegmentType Type { get; set; }
    public int Index { get; set; }
    /// <summary>
    /// 上下文提示（来自前一段的末尾）
    /// </summary>
    public string Context { get; set; }
}

public enum SegmentType
{
    Paragraph,      // 完整段落
    SentenceGroup,  // 句子组
    FixedLength     // 固定长度（最后手段）
}
```

---

### 2. 多策略文本定位（🔴 高）

#### 问题
`FindTextPosition` 使用 Word Find 三级匹配，性能差且容易匹配错误位置。

#### 优化方案
实现多策略定位器：

```csharp
public class SmartTextLocator
{
    /// <summary>
    /// 策略1：使用预建索引（适用于批量定位）
    /// </summary>
    public void BuildDocumentIndex()
    {
        var content = _document.Content.Text;
        var words = ExtractKeyPhrases(content);
        
        foreach (var word in words)
        {
            var positions = FindAllPositions(content, word);
            _indexCache[word] = positions;
        }
    }
    
    /// <summary>
    /// 策略2：模糊匹配（处理 AI 返回文本可能有差异）
    /// </summary>
    public (bool found, int start, int end) FuzzyFindText(
        string target, 
        double similarityThreshold = 0.85)
    {
        // 先尝试精确匹配
        var exact = ExactMatch(target);
        if (exact.found) return exact;
        
        // 使用 Levenshtein 距离计算相似度
        var candidates = FindSimilarTexts(target, similarityThreshold);
        if (candidates.Any())
        {
            var best = candidates.OrderByDescending(c => c.Similarity).First();
            return (true, best.Start, best.End);
        }
        
        return (false, -1, -1);
    }
    
    /// <summary>
    /// 策略3：上下文锚定
    /// </summary>
    public (bool found, int start, int end) ContextualFind(
        string target, 
        string contextBefore, 
        string contextAfter,
        int searchRadius = 100)
    {
        // 先定位上下文
        var (ctxFound, ctxStart, ctxEnd) = FindTextPosition(contextBefore);
        if (!ctxFound) return (false, -1, -1);
        
        // 在上下文范围内搜索目标
        var searchRange = _document.Range(ctxEnd, Math.Min(ctxEnd + searchRadius, _document.Content.End));
        // ... 在范围内搜索 target
    }
}
```

---

### 3. 容错解析器（🔴 高）

#### 问题
`ProofreadIssueParser` 使用单一正则表达式，AI 返回格式不标准时解析失败。

#### 优化方案
实现多模式容错解析：

```csharp
public class RobustProofreadParser
{
    private static readonly List<IParserStrategy> _strategies = new List<IParserStrategy>
    {
        new StandardParserStrategy(),      // 标准格式
        new LooseParserStrategy(),         // 宽松格式
        new MarkdownParserStrategy(),      // Markdown 格式
        new JsonParserStrategy()           // JSON 格式
    };
    
    public List<ProofreadIssueItem> Parse(string aiResponse)
    {
        foreach (var strategy in _strategies.OrderBy(s => s.Priority))
        {
            try
            {
                var result = strategy.Parse(aiResponse);
                if (result.Count > 0)
                {
                    Debug.WriteLine($"使用 {strategy.Name} 解析成功");
                    return result;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"{strategy.Name} 解析失败: {ex.Message}");
            }
        }
        
        // 所有策略失败，返回原始文本
        return new List<ProofreadIssueItem>
        {
            new ProofreadIssueItem
            {
                Type = "原始结果",
                Original = "（解析失败）",
                Modified = aiResponse.Substring(0, Math.Min(200, aiResponse.Length)),
                Reason = "AI返回格式无法识别"
            }
        };
    }
}

/// <summary>
/// 宽松模式解析策略
/// </summary>
public class LooseParserStrategy : IParserStrategy
{
    public string Name => "宽松模式";
    public int Priority => 2;
    
    public List<ProofreadIssueItem> Parse(string text)
    {
        var items = new List<ProofreadIssueItem>();
        
        // 使用更宽松的正则，匹配各种变体
        // 【第X处】或 第X处 或 [X]
        var pattern = @"[【\[]?(?:第)?(?<index>\d+)(?:处)?[】\]]?\s*[：:]\s*(?<type>[^\r\n]+)";
        
        var matches = Regex.Matches(text, pattern, RegexOptions.Multiline);
        
        foreach (Match match in matches)
        {
            var item = new ProofreadIssueItem();
            ParseTypeAndSeverity(match.Groups["type"].Value, out item.Type, out item.Severity);
            
            // 使用关键词查找字段
            var section = ExtractSection(text, match);
            item.Original = ExtractField(section, new[] { "原文", "原文：", "原文:", "原:" });
            item.Modified = ExtractField(section, new[] { "修改", "修改：", "修改:", "改为:" });
            item.Reason = ExtractField(section, new[] { "理由", "理由：", "理由:", "原因", "说明" });
            
            if (!string.IsNullOrEmpty(item.Original))
                items.Add(item);
        }
        
        return items;
    }
}
```

---

### 4. 差异感知增量校验（🟡 中）

#### 问题
`ProofreadIncrementalAsync` 仅简单比较段落哈希，未考虑文本修改类型。

#### 优化方案
使用 Diff 算法实现差异感知：

```csharp
public class DiffAwareIncrementalProofreader
{
    public async Task<List<ParagraphResult>> ProofreadDiffOnly(
        string oldText,
        string newText,
        List<ParagraphResult> previousResults)
    {
        // 1. 使用 Myers Diff 算法找出差异
        var diff = DiffUtil.ComputeDiff(oldText, newText);
        
        var results = new List<ParagraphResult>(previousResults);
        var changedRanges = diff.Where(d => d.Operation != Operation.Equal);
        
        // 2. 只校验变更的段落
        foreach (var change in changedRanges)
        {
            var affectedParagraphs = MapChangeToParagraphs(change, newText);
            
            foreach (var paraIndex in affectedParagraphs)
            {
                var newResult = await ProofreadParagraph(newText, paraIndex);
                results[paraIndex] = newResult;
            }
        }
        
        return results;
    }
    
    public List<int> FindChangedParagraphs(List<string> oldParagraphs, List<string> newParagraphs)
    {
        // 使用最长公共子序列（LCS）算法
        var lcs = LcsAlgorithm.FindLcs(oldParagraphs, newParagraphs);
        var changedIndices = new List<int>();
        
        int oldIdx = 0, newIdx = 0, lcsIdx = 0;
        while (newIdx < newParagraphs.Count)
        {
            if (lcsIdx < lcs.Count && newParagraphs[newIdx] == lcs[lcsIdx])
            {
                oldIdx++; newIdx++; lcsIdx++;
            }
            else
            {
                changedIndices.Add(newIdx);
                newIdx++;
            }
        }
        
        return changedIndices;
    }
}
```

---

### 5. 分层缓存策略（🟡 中）

#### 问题
当前仅有一层内存缓存，且缓存键仅考虑内容哈希。

#### 优化方案
实现多级缓存：

```csharp
public class HierarchicalProofreadCache
{
    // L1: 内存缓存（当前会话）
    private static readonly ConcurrentDictionary<string, CacheEntry> _memoryCache = 
        new ConcurrentDictionary<string, CacheEntry>();
    
    /// <summary>
    /// 生成智能缓存键（考虑内容和位置）
    /// </summary>
    public string GenerateCacheKey(string text, int paragraphIndex, string proofreadMode)
    {
        var contentHash = ComputeHash(text);
        return $"{proofreadMode}:{paragraphIndex}:{contentHash}";
    }
    
    /// <summary>
    /// 相似内容缓存复用（语义缓存）
    /// </summary>
    public bool TryGetSimilarResult(string text, out ParagraphResult result, double threshold = 0.95)
    {
        result = null;
        
        foreach (var entry in _memoryCache.Values)
        {
            var similarity = CalculateSimilarity(entry.Text, text);
            if (similarity >= threshold)
            {
                result = entry.Result;
                result.IsCached = true;
                result.SimilarityScore = similarity;
                return true;
            }
        }
        
        return false;
    }
}
```

---

### 6. 跨段落去重（🟡 中）

#### 问题
重叠段落可能导致同一问题被多次检测。

#### 优化方案
```csharp
public class CrossParagraphDeduplicator
{
    public List<ProofreadIssueItem> Deduplicate(List<ParagraphResult> results)
    {
        var allIssues = results.SelectMany(r => r.Items).ToList();
        var uniqueIssues = new List<ProofreadIssueItem>();
        
        foreach (var issue in allIssues)
        {
            var isDuplicate = uniqueIssues.Any(u => 
                CalculateSimilarity(u.Original, issue.Original) > 0.9 &&
                CalculateSimilarity(u.Modified, issue.Modified) > 0.9);
            
            if (!isDuplicate)
            {
                uniqueIssues.Add(issue);
            }
        }
        
        return uniqueIssues;
    }
}
```

---

## 可靠性优化

### 1. 重试机制（🔴 高）

#### 问题
LLM 请求偶发失败时没有自动重试。

#### 优化方案
在 `BaseLLMService` 中添加指数退避重试：

```csharp
protected virtual async Task<string> PostAsyncWithRetry(
    string jsonContent, 
    RequestLogInfo logInfo, 
    CancellationToken cancellationToken = default, 
    int maxRetries = 3)
{
    for (int i = 0; i < maxRetries; i++)
    {
        try
        {
            return await PostAsync(jsonContent, logInfo, cancellationToken);
        }
        catch (LLMServiceException ex) when (i < maxRetries - 1 && IsRetryableError(ex))
        {
            var delay = TimeSpan.FromSeconds(Math.Pow(2, i)); // 指数退避
            await Task.Delay(delay, cancellationToken);
        }
    }
    throw new LLMServiceException("重试次数已用尽");
}

private bool IsRetryableError(LLMServiceException ex)
{
    return ex.StatusCode == HttpStatusCode.TooManyRequests ||
           ex.StatusCode == HttpStatusCode.ServiceUnavailable ||
           ex.Message.Contains("超时");
}
```

---

### 2. 异常分类体系（🟡 中）

#### 问题
异常处理过于宽泛，难以针对性处理。

#### 优化方案
```csharp
public abstract class ProofreadException : Exception 
{
    protected ProofreadException(string message) : base(message) { }
    protected ProofreadException(string message, Exception inner) : base(message, inner) { }
}

public class DocumentAccessException : ProofreadException 
{
    public DocumentAccessException(string message) : base(message) { }
}

public class LLMConnectionException : ProofreadException 
{
    public LLMConnectionException(string message, Exception inner) : base(message, inner) { }
}

public class CacheException : ProofreadException 
{
    public CacheException(string message) : base(message) { }
}

public class TextLocationException : ProofreadException 
{
    public string TargetText { get; }
    public TextLocationException(string message, string targetText) : base(message) 
    { 
        TargetText = targetText; 
    }
}
```

---

### 3. 请求取消联动（🟡 中）

#### 问题
多个 `CancellationToken` 需要正确联动。

#### 优化方案
确保所有异步操作正确传递和联动 CancellationToken：

```csharp
public async Task<List<ParagraphResult>> ProofreadDocumentAsync(
    string documentText, 
    CancellationToken userCancellationToken = default)
{
    // 创建联动 Token 源
    using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
        userCancellationToken, _disposeCts.Token))
    {
        var linkedToken = linkedCts.Token;
        
        // 所有子任务都使用 linkedToken
        var tasks = paragraphs.Select(p => 
            ProcessParagraphAsync(p, linkedToken)).ToList();
        
        return await Task.WhenAll(tasks);
    }
}
```

---

## 代码质量优化

### 1. 日志框架（🟢 低）

#### 问题
当前使用 `Debug.WriteLine`，生产环境难以收集。

#### 优化方案
引入轻量级日志抽象：

```csharp
public interface ILogger
{
    void Debug(string message);
    void Info(string message);
    void Warning(string message);
    void Error(string message, Exception ex = null);
}

public static class Log
{
    private static ILogger _logger = new DebugLogger();
    
    public static void SetLogger(ILogger logger) => _logger = logger;
    
    public static void Debug(string message) => _logger?.Debug(message);
    public static void Info(string message) => _logger?.Info(message);
    public static void Error(string message, Exception ex = null) => _logger?.Error(message, ex);
}
```

---

### 2. 正则表达式优化（🟢 低）

#### 问题
正则表达式可能性能不佳，且没有超时保护。

#### 优化方案
```csharp
private static readonly Regex ProofreadItemRegex = new Regex(
    @"...pattern...",
    RegexOptions.Singleline | RegexOptions.Compiled | RegexOptions.CultureInvariant,
    TimeSpan.FromSeconds(5)); // 添加超时保护
```

---

### 3. 性能监控（🟢 低）

#### 问题
缺乏性能指标收集。

#### 优化方案
```csharp
public static class PerformanceMetrics
{
    public static long TotalRequests { get; set; }
    public static long TotalRequestTimeMs { get; set; }
    public static long CacheHits { get; set; }
    public static long CacheMisses { get; set; }
    
    public static double AverageRequestTime => 
        TotalRequests > 0 ? (double)TotalRequestTimeMs / TotalRequests : 0;
    
    public static double CacheHitRate => 
        (CacheHits + CacheMisses) > 0 ? (double)CacheHits / (CacheHits + CacheMisses) : 0;
}
```

---

## 安全优化

### 1. 配置热重载（🟢 低）

#### 问题
配置修改需要重启 Word。

#### 优化方案
```csharp
private static FileSystemWatcher _configWatcher;

public static void EnableHotReload()
{
    _configWatcher = new FileSystemWatcher(ConfigDir, "config.dat")
    {
        NotifyFilter = NotifyFilters.LastWrite
    };
    _configWatcher.Changed += (s, e) => 
    {
        // 延迟加载避免文件锁定冲突
        Task.Delay(500).ContinueWith(_ => LoadConfig());
    };
    _configWatcher.EnableRaisingEvents = true;
}
```

---

## 实施路线图

### 第一阶段（立即 - 2 周）
- [ ] 🔴 修复 SHA256 线程安全问题
- [ ] 🔴 修复 HttpClient DNS 刷新问题
- [ ] 🔴 拆分 `GOWordAgentPaneWpf.xaml.cs`
- [ ] 🔴 消除 `CreateIssueButton` 代码重复
- [ ] 🔴 添加 LLM 请求重试机制

### 第二阶段（2-4 周）
- [ ] 🟡 实现智能语义分段
- [ ] 🟡 实现多策略文本定位
- [ ] 🟡 实现容错解析器
- [ ] 🟡 添加缓存内存上限
- [ ] 🟡 实现异常分类体系

### 第三阶段（4-6 周）
- [ ] 🟢 引入日志框架
- [ ] 🟢 实现差异感知增量校验
- [ ] 🟢 实现分层缓存策略
- [ ] 🟢 添加性能监控
- [ ] 🟢 单元测试覆盖

### 第四阶段（可选）
- [ ] 🟢 流式响应支持
- [ ] 🟢 配置热重载
- [ ] 🟢 UI 主题自定义

---

## 附录

### A. 检查清单

#### 提交前检查
- [ ] 代码是否遵循现有命名规范
- [ ] 是否添加了必要的异常处理
- [ ] 是否更新了相关文档
- [ ] 是否测试了 Word 2016/2019/365 兼容性
- [ ] 是否检查了 COM 对象释放

#### 性能检查
- [ ] 是否避免在循环中创建对象
- [ ] 是否正确使用 `ConfigureAwait(false)`
- [ ] 是否添加了 `RegexOptions.Compiled`
- [ ] 是否考虑了大数据量下的内存占用

### B. 参考资源

- [VSTO 开发最佳实践](https://docs.microsoft.com/zh-cn/visualstudio/vsto/office-solutions-development-overview-vsto)
- [WPF 性能优化](https://docs.microsoft.com/zh-cn/dotnet/desktop/wpf/advanced/optimizing-wpf-application-performance)
- [HttpClient 最佳实践](https://docs.microsoft.com/zh-cn/dotnet/fundamentals/networking/http/httpclient-guidelines)

---

*文档结束*
