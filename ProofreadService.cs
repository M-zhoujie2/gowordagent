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
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{

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
                            Items = prev.Items?.ToList() ?? new List<ProofreadIssueItem>()
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
                string response;
                try
                {
                    response = await _llmService.SendProofreadMessageAsync(_systemPrompt, paragraph);
                }
                catch (LLMServiceException ex)
                {
                    // API 调用失败，记录错误但继续处理其他段落
                    Debug.WriteLine($"[ProofreadService] 段落 {index + 1} LLM调用失败: {ex.GetFriendlyErrorMessage()}");
                    
                    return new ParagraphResult
                    {
                        Index = index,
                        OriginalText = paragraph,
                        ResultText = $"[错误] {ex.GetFriendlyErrorMessage()}",
                        IsCompleted = true,
                        IsCached = false,
                        ProcessTime = DateTime.Now,
                        ElapsedMs = stopwatch.ElapsedMilliseconds,
                        Items = new System.Collections.Generic.List<ProofreadIssueItem>()
                    };
                }
                
                Debug.WriteLine($"[ProofreadService] 段落 {index + 1} LLM返回, 结果长度={response?.Length ?? 0}");
                
                // 解析结果
                var items = ProofreadIssueParser.ParseProofreadItems(response);
                
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
        /// 生成统计报告（详细版）
        /// </summary>
        /// <param name="results">段落结果列表</param>
        /// <param name="totalChars">总字数（可选）</param>
        /// <param name="elapsed">耗时（可选）</param>
        /// <param name="providerName">AI 提供商名称（可选）</param>
        public static string GenerateReport(List<ParagraphResult> results, int totalChars = 0, TimeSpan? elapsed = null, string providerName = null)
        {
            var sb = new StringBuilder();
            var now = DateTime.Now;
            
            sb.AppendLine("【校对报告】");
            sb.AppendLine();
            sb.AppendLine("基本信息：");
            
            if (totalChars > 0)
                sb.AppendLine($"  字数：{totalChars:N0}");
            
            sb.AppendLine($"  分块：{results.Count} 块");
            
            if (!string.IsNullOrEmpty(providerName))
                sb.AppendLine($"  校对模型：{providerName}");
            
            sb.AppendLine($"  生成时间：{now:yyyy-MM-dd HH:mm:ss}");
            
            if (elapsed.HasValue)
                sb.AppendLine($"  耗时：{elapsed.Value.TotalSeconds:F1} 秒");
            
            sb.AppendLine();

            int totalIssues = 0;
            var allCategories = new Dictionary<string, int>();
            int cachedCount = 0;

            foreach (var result in results)
            {
                if (result == null) continue;
                
                int issues = ProofreadIssueParser.CountIssues(result.ResultText);
                totalIssues += issues;

                var cats = ProofreadIssueParser.CategorizeIssues(result.ResultText);
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

            sb.AppendLine($"统计汇总（共发现 {totalIssues} 处问题）：");
            
            if (allCategories.Count > 0)
            {
                foreach (var kv in allCategories.OrderByDescending(x => x.Value))
                {
                    sb.AppendLine($"  • {kv.Key}：{kv.Value} 处");
                }
            }
            else
            {
                sb.AppendLine("  • 未发现明显错误");
            }
            
            if (cachedCount > 0)
            {
                sb.AppendLine();
                sb.AppendLine($"（其中 {cachedCount} 段来自缓存）");
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
}
