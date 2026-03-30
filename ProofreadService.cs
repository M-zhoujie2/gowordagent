using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace GOWordAgentAddIn
{
    #region 数据模型

    /// <summary>
    /// 段落校验结果
    /// </summary>
    public class ParagraphResult
    {
        /// <summary>
        /// 段落索引
        /// </summary>
        public int Index { get; set; }
        
        /// <summary>
        /// 原始文本
        /// </summary>
        public string OriginalText { get; set; }
        
        /// <summary>
        /// 完整结果文本
        /// </summary>
        public string ResultText { get; set; }
        
        /// <summary>
        /// 是否已完成
        /// </summary>
        public bool IsCompleted { get; set; }
        
        /// <summary>
        /// 是否来自缓存
        /// </summary>
        public bool IsCached { get; set; }
        
        /// <summary>
        /// 处理时间
        /// </summary>
        public DateTime ProcessTime { get; set; }
        
        /// <summary>
        /// 耗时（毫秒）
        /// </summary>
        public long ElapsedMs { get; set; }
        
        /// <summary>
        /// 解析出的校对项
        /// </summary>
        public List<ProofreadItem> Items { get; set; } = new List<ProofreadItem>();
    }

    /// <summary>
    /// 校对项
    /// </summary>
    public class ProofreadItem
    {
        public string Type { get; set; }
        public string Severity { get; set; }
        public string Original { get; set; }
        public string Modified { get; set; }
        public string Reason { get; set; }
    }

    #endregion

    #region 事件参数

    /// <summary>
    /// 校验进度回调参数
    /// </summary>
    public class ProofreadProgressArgs : EventArgs
    {
        /// <summary>
        /// 总段落数
        /// </summary>
        public int TotalParagraphs { get; set; }
        
        /// <summary>
        /// 已完成段落数
        /// </summary>
        public int CompletedParagraphs { get; set; }
        
        /// <summary>
        /// 当前处理中的段落索引
        /// </summary>
        public int CurrentIndex { get; set; }
        
        /// <summary>
        /// 当前状态描述
        /// </summary>
        public string CurrentStatus { get; set; }
        
        /// <summary>
        /// 当前段落结果
        /// </summary>
        public ParagraphResult Result { get; set; }
        
        /// <summary>
        /// 是否全部完成
        /// </summary>
        public bool IsCompleted { get; set; }
        
        /// <summary>
        /// 预计剩余时间（秒）
        /// </summary>
        public int EstimatedRemainingSeconds { get; set; }
        
        /// <summary>
        /// 缓存命中数
        /// </summary>
        public int CacheHitCount { get; set; }
    }

    #endregion

    #region 校对服务

    /// <summary>
    /// 增强的校验服务 - 支持并发、缓存
    /// </summary>
    public class ProofreadService : IDisposable
    {
        private readonly ILLMService _llmService;
        private readonly string _systemPrompt;
        private readonly int _concurrency;
        private readonly SemaphoreSlim _semaphore;
        private readonly Dispatcher _dispatcher;

        // 文档分段器
        private readonly DocumentSegmenter _segmenter;

        // 性能统计 - 使用 Interlocked 操作
        private int _cacheHitCount = 0;

        // 预编译的正则表达式（提高性能）
        private static readonly Regex ProofreadItemRegex = new Regex(
            @"【第\d+处】类型：(?<type>[^\r\n|]+)(?:[｜|]严重度：(?<severity>[^\r\n]+))?\r?\n原文：(?<original>.*?)\r?\n修改：(?<modified>.*?)\r?\n理由：(?<reason>.*?)(?=\r?\n【第|$)", 
            RegexOptions.Singleline | RegexOptions.Compiled);

        /// <summary>
        /// 进度更新事件（线程安全）
        /// </summary>
        public event EventHandler<ProofreadProgressArgs> OnProgress
        {
            add { _onProgress += value; }
            remove { _onProgress -= value; }
        }
        private EventHandler<ProofreadProgressArgs> _onProgress;

        /// <summary>
        /// 创建校对服务实例
        /// </summary>
        /// <param name="llmService">LLM服务</param>
        /// <param name="systemPrompt">系统提示词</param>
        /// <param name="concurrency">并发数（1-10）</param>
        public ProofreadService(ILLMService llmService, string systemPrompt, int concurrency = 5, SegmenterConfig segmenterConfig = null)
        {
            _llmService = llmService ?? throw new ArgumentNullException(nameof(llmService));
            _systemPrompt = systemPrompt ?? throw new ArgumentNullException(nameof(systemPrompt));
            _concurrency = Math.Max(1, Math.Min(concurrency, 10)); // 限制1-10并发
            _semaphore = new SemaphoreSlim(_concurrency);
            _dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
            _segmenter = new DocumentSegmenter(segmenterConfig);
        }

        #region 公共方法

        /// <summary>
        /// 开始校验文档 - 优化版：最先完成先显示
        /// </summary>
        public async Task<List<ParagraphResult>> ProofreadDocumentAsync(string documentText, CancellationToken cancellationToken = default)
        {
            Debug.WriteLine($"[ProofreadDocumentAsync] 开始文档校对，文档长度={documentText?.Length ?? 0}");
            
            // 1. 分段
            var paragraphs = _segmenter.SplitIntoParagraphs(documentText);
            var totalCount = paragraphs.Count;
            
            if (totalCount == 0)
            {
                return new List<ParagraphResult>();
            }

            // 2. 创建结果容器
            var results = new List<ParagraphResult>(new ParagraphResult[totalCount]); // 预分配固定位置
            var completedCount = 0;
            var lockObj = new object();
            var stopwatch = Stopwatch.StartNew();

            // 3. 启动所有任务（受信号量控制并发）
            var pendingTasks = new List<Task<ParagraphResult>>();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var index = i; // 捕获索引
                var para = paragraphs[i];
                
                var task = Task.Run(async () => 
                {
                    return await ProcessParagraphAsync(para, index, totalCount, cancellationToken);
                }, cancellationToken);
                
                pendingTasks.Add(task);
            }

            // 4. 使用 Task.WhenAny 实现"最先完成先显示"
            while (pendingTasks.Count > 0)
            {
                // 等待任意一个任务完成
                var completedTask = await Task.WhenAny(pendingTasks).ConfigureAwait(false);
                pendingTasks.Remove(completedTask);
                
                var result = await completedTask;
                
                // 将结果放入正确位置
                lock (lockObj)
                {
                    results[result.Index] = result;
                    completedCount++;
                    
                    if (result.IsCached)
                        System.Threading.Interlocked.Increment(ref _cacheHitCount);
                }
                
                // 计算预计剩余时间
                var estimatedRemaining = CalculateEstimatedRemaining(completedCount, totalCount, stopwatch.Elapsed);
                
                // 立即报告进度（最先完成先显示）
                try
                {
                    await _dispatcher.InvokeAsync(() =>
                        ReportProgress(totalCount, completedCount, result.Index, 
                            result.IsCached ? $"第 {result.Index + 1} 段（缓存）" : $"第 {result.Index + 1} 段完成", 
                            result, false, estimatedRemaining));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                }
            }

            stopwatch.Stop();

            // 5. 报告完成
            try
            {
                await _dispatcher.InvokeAsync(() => 
                    ReportProgress(totalCount, totalCount, -1, 
                        $"校对完成（耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒）", 
                        null, true, 0));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProofreadService] 报告完成时出错: {ex.Message}");
            }

            Debug.WriteLine($"[ProofreadDocumentAsync] 校对完成，总耗时={stopwatch.Elapsed.TotalSeconds:F1}秒");

            return results.ToList();
        }

        /// <summary>
        /// 增量校验 - 只校验修改的段落
        /// </summary>
        public async Task<List<ParagraphResult>> ProofreadIncrementalAsync(
            string documentText, 
            List<ParagraphResult> previousResults,
            CancellationToken cancellationToken = default)
        {
            Debug.WriteLine($"[ProofreadIncrementalAsync] 开始增量校对");
            
            var paragraphs = _segmenter.SplitIntoParagraphs(documentText);
            var totalCount = paragraphs.Count;
            
            if (totalCount == 0)
                return new List<ParagraphResult>();

            var results = new List<ParagraphResult>(new ParagraphResult[totalCount]);
            var tasks = new List<Task<ParagraphResult>>();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var index = i;
                var para = paragraphs[i];
                var cacheKey = ProofreadCacheManager.ComputeHash(para);

                // 检查是否与之前相同
                if (previousResults != null && index < previousResults.Count)
                {
                    var prev = previousResults[index];
                    if (prev != null && ProofreadCacheManager.ComputeHash(prev.OriginalText) == cacheKey)
                    {
                        // 段落未变化，复用结果
                        results[index] = new ParagraphResult
                        {
                            Index = index,
                            OriginalText = para,
                            ResultText = prev.ResultText,
                            IsCompleted = true,
                            IsCached = true,
                            ProcessTime = DateTime.Now,
                            Items = prev.Items?.ToList() ?? new List<ProofreadItem>()
                        };
                        continue;
                    }
                }

                // 需要重新校验
                var task = Task.Run(async () => 
                {
                    return await ProcessParagraphAsync(para, index, totalCount, cancellationToken);
                }, cancellationToken);
                
                tasks.Add(task);
            }

            // 等待所有新任务完成
            while (tasks.Count > 0)
            {
                var completedTask = await Task.WhenAny(tasks);
                tasks.Remove(completedTask);
                var result = await completedTask;
                results[result.Index] = result;
            }

            return results.Where(r => r != null).ToList();
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 处理单个段落（带缓存检查）
        /// </summary>
        private async Task<ParagraphResult> ProcessParagraphAsync(string paragraph, int index, int total, 
            CancellationToken cancellationToken)
        {
            await _semaphore.WaitAsync(cancellationToken);
            var stopwatch = Stopwatch.StartNew();
            
            try
            {
                Debug.WriteLine($"[ProofreadService] 开始处理段落 {index + 1}/{total}, 长度={paragraph.Length}");
                
                // 检查缓存
                if (ProofreadCacheManager.TryGetCachedResult(paragraph, index, out var cachedResult))
                {
                    stopwatch.Stop();
                    return cachedResult;
                }

                // 报告开始处理
                try
                {
                    await _dispatcher.InvokeAsync(() => 
                        ReportProgress(total, index, index, $"正在校对第 {index + 1}/{total} 段...", 
                            null, false, -1));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                }

                // 调用LLM
                Debug.WriteLine($"[ProofreadService] 段落 {index + 1} 调用LLM...");
                var response = await _llmService.SendProofreadMessageAsync(_systemPrompt, paragraph);
                Debug.WriteLine($"[ProofreadService] 段落 {index + 1} LLM返回, 结果长度={response?.Length ?? 0}");
                
                // 解析结果
                var items = ParseProofreadItems(response);
                
                var result = new ParagraphResult
                {
                    Index = index,
                    OriginalText = paragraph,
                    ResultText = response,
                    IsCompleted = true,
                    IsCached = false,
                    ProcessTime = DateTime.Now,
                    ElapsedMs = stopwatch.ElapsedMilliseconds,
                    Items = items
                };

                stopwatch.Stop();

                // 存入缓存
                ProofreadCacheManager.StoreResult(paragraph, result);

                return result;
            }
            finally
            {
                _semaphore.Release();
            }
        }

        /// <summary>
        /// 计算预计剩余时间
        /// </summary>
        private int CalculateEstimatedRemaining(int completed, int total, TimeSpan elapsed)
        {
            if (completed == 0 || total == 0) return -1;
            
            var avgTimePerItem = elapsed.TotalSeconds / completed;
            var remaining = total - completed;
            return (int)(avgTimePerItem * remaining);
        }

        /// <summary>
        /// 解析校对结果
        /// </summary>
        private List<ProofreadItem> ParseProofreadItems(string aiResponse)
        {
            var items = new List<ProofreadItem>();
            
            if (string.IsNullOrWhiteSpace(aiResponse))
                return items;
            
            // 使用预编译的正则表达式（静态字段，性能更好）
            foreach (Match match in ProofreadItemRegex.Matches(aiResponse))
            {
                items.Add(new ProofreadItem
                {
                    Type = match.Groups["type"].Value.Trim(),
                    Severity = match.Groups["severity"].Value.Trim().ToLower(),
                    Original = match.Groups["original"].Value.Trim(),
                    Modified = match.Groups["modified"].Value.Trim(),
                    Reason = match.Groups["reason"].Value.Trim()
                });
            }

            return items;
        }

        /// <summary>
        /// 报告进度
        /// </summary>
        private void ReportProgress(int total, int completed, int current, string status, 
            ParagraphResult result, bool isCompleted, int estimatedRemaining = -1)
        {
            // 创建本地副本避免竞态条件
            var handler = _onProgress;
            handler?.Invoke(this, new ProofreadProgressArgs
            {
                TotalParagraphs = total,
                CompletedParagraphs = completed,
                CurrentIndex = current,
                CurrentStatus = status,
                Result = result,
                IsCompleted = isCompleted,
                EstimatedRemainingSeconds = estimatedRemaining,
                CacheHitCount = System.Threading.Interlocked.CompareExchange(ref _cacheHitCount, 0, 0)
            });
        }

        #endregion

        #region 静态工具方法

        /// <summary>
        /// 统计问题数量
        /// </summary>
        public static int CountIssues(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var matches = Regex.Matches(text, @"【第\d+处】");
            return matches.Count;
        }

        /// <summary>
        /// 按类型统计问题
        /// </summary>
        public static Dictionary<string, int> CategorizeIssues(string text)
        {
            var categories = new Dictionary<string, int>();
            if (string.IsNullOrWhiteSpace(text)) return categories;

            var pattern = @"【第\d+处】类型：([^\r\n|]+)";
            foreach (Match match in Regex.Matches(text, pattern))
            {
                string cat = match.Groups[1].Value.Trim();
                
                // 归一化分类
                if (cat.Contains("错别字") || cat.Contains("拼写"))
                    cat = "错别字";
                else if (cat.Contains("语病") || cat.Contains("语法") || cat.Contains("搭配") || cat.Contains("杂糅"))
                    cat = "语病";
                else if (cat.Contains("标点"))
                    cat = "标点错误";
                else if (cat.Contains("序号"))
                    cat = "序号问题";
                else if (cat.Contains("用词"))
                    cat = "用词不当";
                else if (cat.Contains("术语") || cat.Contains("不一致"))
                    cat = "术语不一致";
                else if (cat.Contains("格式"))
                    cat = "格式问题";
                else if (cat.Contains("逻辑") || cat.Contains("矛盾"))
                    cat = "逻辑/矛盾";
                else
                    cat = "其他";

                if (categories.ContainsKey(cat))
                    categories[cat]++;
                else
                    categories[cat] = 1;
            }

            return categories;
        }

        /// <summary>
        /// 生成统计报告
        /// </summary>
        public static string GenerateReport(List<ParagraphResult> results)
        {
            var sb = new StringBuilder();
            sb.AppendLine("## 校对统计报告");
            sb.AppendLine();

            int totalIssues = 0;
            var allCategories = new Dictionary<string, int>();
            int cachedCount = 0;

            foreach (var result in results)
            {
                if (result == null) continue;
                
                int issues = CountIssues(result.ResultText);
                totalIssues += issues;

                var cats = CategorizeIssues(result.ResultText);
                foreach (var kv in cats)
                {
                    if (allCategories.ContainsKey(kv.Key))
                        allCategories[kv.Key] += kv.Value;
                    else
                        allCategories[kv.Key] = kv.Value;
                }

                if (result.IsCached)
                    cachedCount++;
            }

            sb.AppendLine($"**共处理 {results.Count} 段，发现 {totalIssues} 处问题**");
            if (cachedCount > 0)
                sb.AppendLine($"（其中 {cachedCount} 段来自缓存）");
            sb.AppendLine();

            if (allCategories.Count > 0)
            {
                sb.AppendLine("### 问题分类统计");
                foreach (var kv in allCategories.OrderByDescending(x => x.Value))
                {
                    sb.AppendLine($"- {kv.Key}：{kv.Value} 处");
                }
            }
            else
            {
                sb.AppendLine("✅ 未发现明显错误");
            }

            return sb.ToString();
        }

        /// <summary>
        /// 清除缓存
        /// </summary>
        public static void ClearCache()
        {
            ProofreadCacheManager.ClearCache();
        }

        /// <summary>
        /// 获取缓存统计
        /// </summary>
        public static (int count, long totalBytes) GetCacheStats()
        {
            return ProofreadCacheManager.GetCacheStats();
        }

        #endregion

        #region IDisposable

        private volatile bool _disposed = false;

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源的实际实现
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 释放托管资源
                    _semaphore?.Dispose();
                }
                
                _disposed = true;
            }
        }

        /// <summary>
        /// 终结器
        /// </summary>
        ~ProofreadService()
        {
            Dispose(false);
        }

        #endregion
    }

    #endregion
}
