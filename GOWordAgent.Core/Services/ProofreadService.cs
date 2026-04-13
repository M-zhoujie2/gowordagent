using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对服务 - 支持并发、缓存
    /// </summary>
    public class ProofreadService : IDisposable
    {
        private readonly ILLMService _llmService;
        private readonly string _systemPrompt;
        private readonly int _concurrency;
        private readonly SemaphoreSlim _semaphore;
        private readonly string _proofreadMode;
        private readonly DocumentSegmenter _segmenter;
        private int _cacheHitCount = 0;
        private readonly CancellationTokenSource _disposeCts = new CancellationTokenSource();
        private long _activeTaskCount = 0;

        public event EventHandler<ProofreadProgressArgs>? OnProgress;

        public ProofreadService(ILLMService llmService, string systemPrompt, int concurrency = 5, SegmenterConfig? segmenterConfig = null, string? proofreadMode = null)
        {
            _llmService = llmService ?? throw new ArgumentNullException(nameof(llmService));
            _systemPrompt = systemPrompt ?? throw new ArgumentNullException(nameof(systemPrompt));
            _concurrency = Math.Max(1, Math.Min(concurrency, 10));
            _semaphore = new SemaphoreSlim(_concurrency);
            _segmenter = new DocumentSegmenter(segmenterConfig);
            _proofreadMode = proofreadMode ?? "精准校验";
        }

        public async Task<List<ParagraphResult>> ProofreadDocumentAsync(string documentText, CancellationToken cancellationToken = default)
        {
            Debug.WriteLine($"[ProofreadDocumentAsync] 开始文档校对，文档长度={documentText?.Length ?? 0}");

            var paragraphs = _segmenter.SplitIntoParagraphs(documentText);
            var totalCount = paragraphs.Count;

            if (totalCount == 0)
            {
                return new List<ParagraphResult>();
            }

            var results = new List<ParagraphResult>(new ParagraphResult[totalCount]);
            var completedCount = 0;
            var lockObj = new object();
            var stopwatch = Stopwatch.StartNew();

            var pendingTasks = new List<Task<ParagraphResult>>();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var index = i;
                var para = paragraphs[i];

                var task = Task.Run(async () =>
                {
                    return await ProcessParagraphAsync(para, index, totalCount, cancellationToken);
                }, cancellationToken);

                pendingTasks.Add(task);
            }

            while (pendingTasks.Count > 0)
            {
                var completedTask = await Task.WhenAny(pendingTasks).ConfigureAwait(false);
                pendingTasks.Remove(completedTask);

                var result = await completedTask;

                lock (lockObj)
                {
                    results[result.Index] = result;
                    completedCount++;

                    if (result.IsCached)
                        System.Threading.Interlocked.Increment(ref _cacheHitCount);
                }

                var estimatedRemaining = CalculateEstimatedRemaining(completedCount, totalCount, stopwatch.Elapsed);

                try
                {
                    ReportProgress(totalCount, completedCount, result.Index,
                        result.IsCached ? $"第 {result.Index + 1} 段（缓存）" : $"第 {result.Index + 1} 段完成",
                        result, false, estimatedRemaining);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                }
            }

            stopwatch.Stop();

            try
            {
                ReportProgress(totalCount, totalCount, -1,
                    $"校对完成（耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒）",
                    null, true, 0);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProofreadService] 报告完成时出错: {ex.Message}");
            }

            Debug.WriteLine($"[ProofreadDocumentAsync] 校对完成，总耗时={stopwatch.Elapsed.TotalSeconds:F1}秒");

            return results.ToList();
        }

        private async Task<ParagraphResult> ProcessParagraphAsync(string paragraph, int index, int total,
            CancellationToken cancellationToken)
        {
            using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _disposeCts.Token))
            {
                var linkedToken = linkedCts.Token;
                await _semaphore.WaitAsync(linkedToken);

                System.Threading.Interlocked.Increment(ref _activeTaskCount);
                var stopwatch = Stopwatch.StartNew();

                try
                {
                    Debug.WriteLine($"[ProofreadService] 开始处理段落 {index + 1}/{total}, 长度={paragraph.Length}");

                    if (ProofreadCacheManager.TryGetCachedResult(paragraph, index, out var cachedResult, _proofreadMode))
                    {
                        stopwatch.Stop();
                        return cachedResult!;
                    }

                    try
                    {
                        ReportProgress(total, index, index, $"正在校对第 {index + 1}/{total} 段...", null, false, -1);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                    }

                    Debug.WriteLine($"[ProofreadService] 段落 {index + 1} 调用LLM...");
                    var response = await _llmService.SendProofreadMessageAsync(_systemPrompt, paragraph);
                    Debug.WriteLine($"[ProofreadService] 段落 {index + 1} LLM返回, 结果长度={response?.Length ?? 0}");

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

                    ProofreadCacheManager.StoreResult(paragraph, result, _proofreadMode);

                    return result;
                }
                finally
                {
                    _semaphore.Release();
                    System.Threading.Interlocked.Decrement(ref _activeTaskCount);
                }
            }
        }

        private int CalculateEstimatedRemaining(int completed, int total, TimeSpan elapsed)
        {
            if (completed == 0 || total == 0) return -1;

            var avgTimePerItem = elapsed.TotalSeconds / completed;
            var remaining = total - completed;
            return (int)(avgTimePerItem * remaining);
        }

        private void ReportProgress(int total, int completed, int current, string status,
            ParagraphResult? result, bool isCompleted, int estimatedRemaining = -1)
        {
            var handler = OnProgress;
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

        public static string GenerateReport(List<ParagraphResult> results, int totalChars = 0, TimeSpan? elapsed = null, string? providerName = null)
        {
            var sb = new StringBuilder();
            var now = DateTime.Now;

            sb.AppendLine("# 校对报告");
            sb.AppendLine();
            sb.AppendLine("## 基本信息");

            if (totalChars > 0)
                sb.AppendLine($"- **字数**：{totalChars:N0}");

            sb.AppendLine($"- **分块**：{results.Count} 块");

            if (!string.IsNullOrEmpty(providerName))
                sb.AppendLine($"- **校对模型**：{providerName}");

            sb.AppendLine($"- **生成时间**：{now:yyyy-MM-dd HH:mm:ss}");

            if (elapsed.HasValue)
                sb.AppendLine($"- **耗时**：{elapsed.Value.TotalSeconds:F1} 秒");

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

            sb.AppendLine("## 统计汇总");
            sb.AppendLine();
            sb.AppendLine($"### 发现问题（共 {totalIssues} 处）");

            if (allCategories.Count > 0)
            {
                foreach (var kv in allCategories.OrderByDescending(x => x.Value))
                {
                    sb.AppendLine($"- {kv.Key}：{kv.Value} 处");
                }
            }
            else
            {
                sb.AppendLine("- 未发现明显错误");
            }

            if (cachedCount > 0)
            {
                sb.AppendLine();
                sb.AppendLine($"（其中 {cachedCount} 段来自缓存）");
            }

            return sb.ToString();
        }

        public static void ClearCache()
        {
            ProofreadCacheManager.ClearCache();
        }

        public static (int count, long totalBytes) GetCacheStats()
        {
            return ProofreadCacheManager.GetCacheStats();
        }

        #region IDisposable

        private volatile bool _disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _disposeCts?.Cancel();

                    var sw = Stopwatch.StartNew();
                    while (sw.ElapsedMilliseconds < 5000 && System.Threading.Interlocked.Read(ref _activeTaskCount) > 0)
                    {
                        Thread.Sleep(50);
                    }

                    _semaphore?.Dispose();
                    _disposeCts?.Dispose();
                }

                _disposed = true;
            }
        }

        #endregion
    }
}
