using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对服务 - 支持并发、缓存、部分结果保留和全局取消
    /// </summary>
    public class ProofreadService : IDisposable
    {
        private readonly ILLMService _llmService;
        private readonly string _systemPrompt;
        private readonly int _concurrency;
        private readonly string _proofreadMode;
        private readonly DocumentSegmenter _segmenter;
        private int _cacheHitCount = 0;
        private readonly CancellationTokenSource _disposeCts = new CancellationTokenSource();

        // 全局取消：用于服务端主动中断所有正在进行的校对（单用户本地服务场景）
        private static CancellationTokenSource _globalCancelCts = new CancellationTokenSource();
        private static readonly object _globalCancelLock = new object();

        // 进度节流
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private readonly TimeSpan _progressThrottleInterval = TimeSpan.FromMilliseconds(200);
        private readonly object _progressLock = new object();

        public event EventHandler<ProofreadProgressArgs>? OnProgress;

        public ProofreadService(ILLMService llmService, string systemPrompt, int concurrency = 5, SegmenterConfig? segmenterConfig = null, string? proofreadMode = null)
        {
            _llmService = llmService ?? throw new ArgumentNullException(nameof(llmService));
            _systemPrompt = systemPrompt ?? throw new ArgumentNullException(nameof(systemPrompt));
            _concurrency = Math.Max(1, Math.Min(concurrency, 10));
            _segmenter = new DocumentSegmenter(segmenterConfig);
            _proofreadMode = proofreadMode ?? "精准校验";
        }

        public static void CancelAll()
        {
            lock (_globalCancelLock)
            {
                if (!_globalCancelCts.IsCancellationRequested)
                {
                    try { _globalCancelCts.Cancel(); }
                    catch (ObjectDisposedException) { }
                }
            }
        }

        public static void ResetGlobalCancellation()
        {
            lock (_globalCancelLock)
            {
                if (_globalCancelCts.IsCancellationRequested)
                {
                    var old = _globalCancelCts;
                    _globalCancelCts = new CancellationTokenSource();
                    try { old.Dispose(); }
                    catch { }
                }
            }
        }

        public async Task<List<ParagraphResult>> ProofreadDocumentAsync(string? documentText, CancellationToken cancellationToken = default)
        {
            Debug.WriteLine($"[ProofreadDocumentAsync] 开始文档校对，文档长度={documentText?.Length ?? 0}");

            var paragraphs = _segmenter.SplitIntoParagraphs(documentText ?? "");
            var totalCount = paragraphs.Count;

            if (totalCount == 0)
            {
                return new List<ParagraphResult>();
            }

            var results = new ParagraphResult[totalCount];
            var completedCount = 0;
            var failedCount = 0;
            var lockObj = new object();
            var stopwatch = Stopwatch.StartNew();

            var indexedParagraphs = paragraphs.Select((text, i) => (Text: text, Index: i)).ToList();

            CancellationTokenSource? requestCts = null;
            try
            {
                // 链接三个取消源：HTTP 请求取消、服务实例 Dispose、全局取消
                requestCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken, _disposeCts.Token, _globalCancelCts.Token);

                await Parallel.ForEachAsync(
                    indexedParagraphs,
                    new ParallelOptions
                    {
                        MaxDegreeOfParallelism = _concurrency,
                        CancellationToken = requestCts.Token
                    },
                    async (item, ct) =>
                    {
                        ParagraphResult result;
                        try
                        {
                            result = await ProcessParagraphAsync(item.Text, item.Index, totalCount, ct).ConfigureAwait(false);
                        }
                        catch (OperationCanceledException)
                        {
                            throw; // 取消需要冒泡，避免继续浪费资源
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ProofreadService] 段落 {item.Index + 1} 处理失败: {ex.Message}");
                            result = new ParagraphResult
                            {
                                Index = item.Index,
                                OriginalText = item.Text,
                                ResultText = $"处理失败: {ex.Message}",
                                IsCompleted = false,
                                IsCached = false,
                                ProcessTime = DateTime.UtcNow,
                                ElapsedMs = 0,
                                Items = new List<ProofreadIssueItem>()
                            };
                        }

                        lock (lockObj)
                        {
                            results[result.Index] = result;
                            if (result.IsCompleted)
                            {
                                completedCount++;
                                if (result.IsCached)
                                    System.Threading.Interlocked.Increment(ref _cacheHitCount);
                            }
                            else
                            {
                                failedCount++;
                            }
                        }

                        var estimatedRemaining = CalculateEstimatedRemaining(
                            completedCount + failedCount, totalCount, stopwatch.Elapsed);

                        try
                        {
                            var statusMsg = result.IsCompleted
                                ? (result.IsCached ? $"第 {result.Index + 1} 段（缓存）" : $"第 {result.Index + 1} 段完成")
                                : $"第 {result.Index + 1} 段失败";
                            ReportProgressThrottled(totalCount, completedCount + failedCount, result.Index,
                                statusMsg, result, false, estimatedRemaining);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                        }
                    }).ConfigureAwait(false);
            }
            finally
            {
                requestCts?.Dispose();
            }

            stopwatch.Stop();

            try
            {
                var finalStatus = failedCount > 0
                    ? $"校对完成（耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒，{failedCount} 段失败）"
                    : $"校对完成（耗时 {stopwatch.Elapsed.TotalSeconds:F1} 秒）";
                ReportProgress(totalCount, totalCount, -1, finalStatus, null, true, 0);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProofreadService] 报告完成时出错: {ex.Message}");
            }

            Debug.WriteLine($"[ProofreadDocumentAsync] 校对完成，总耗时={stopwatch.Elapsed.TotalSeconds:F1}秒，失败段落={failedCount}");

            return results.ToList();
        }

        private async Task<ParagraphResult> ProcessParagraphAsync(string paragraph, int index, int total,
            CancellationToken cancellationToken)
        {
            // 检查是否已释放
            if (_disposed)
                throw new ObjectDisposedException(nameof(ProofreadService));

            using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _disposeCts.Token))
            {
                var linkedToken = linkedCts.Token;
                var stopwatch = Stopwatch.StartNew();

                Debug.WriteLine($"[ProofreadService] 开始处理段落 {index + 1}/{total}, 长度={paragraph.Length}");

                if (ProofreadCacheManager.TryGetCachedResult(paragraph, index, out var cachedResult, _proofreadMode))
                {
                    stopwatch.Stop();
                    return cachedResult!;
                }

                try
                {
                    // 使用节流的进度报告
                    ReportProgressThrottled(total, index, index, $"正在校对第 {index + 1}/{total} 段...", null, false, -1);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ProofreadService] 报告进度时出错: {ex.Message}");
                }

                Debug.WriteLine($"[ProofreadService] 段落 {index + 1} 调用LLM...");
                var response = await _llmService.SendProofreadMessageAsync(_systemPrompt, paragraph, linkedToken);
                Debug.WriteLine($"[ProofreadService] 段落 {index + 1} LLM返回, 结果长度={response?.Length ?? 0}");

                var items = ProofreadIssueParser.ParseProofreadItems(response ?? "");

                var result = new ParagraphResult
                {
                    Index = index,
                    OriginalText = paragraph,
                    ResultText = response ?? "",
                    IsCompleted = true,
                    IsCached = false,
                    ProcessTime = DateTime.UtcNow,
                    ElapsedMs = stopwatch.ElapsedMilliseconds,
                    Items = items
                };

                stopwatch.Stop();

                ProofreadCacheManager.StoreResult(paragraph, result, _proofreadMode);

                return result;
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

        /// <summary>
        /// 节流的进度报告，避免过于频繁的 UI 更新
        /// </summary>
        private void ReportProgressThrottled(int total, int completed, int current, string status,
            ParagraphResult? result, bool isCompleted, int estimatedRemaining = -1)
        {
            // 完成状态或错误状态总是报告
            if (isCompleted || status.Contains("错误") || status.Contains("失败"))
            {
                lock (_progressLock)
                {
                    _lastProgressUpdate = DateTime.Now;
                }
                ReportProgress(total, completed, current, status, result, isCompleted, estimatedRemaining);
                return;
            }

            // 检查是否需要节流
            lock (_progressLock)
            {
                var now = DateTime.Now;
                if (now - _lastProgressUpdate < _progressThrottleInterval)
                {
                    return; // 跳过这次更新
                }
                _lastProgressUpdate = now;
            }

            ReportProgress(total, completed, current, status, result, isCompleted, estimatedRemaining);
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

        /// <summary>
        /// 请求取消所有正在进行的任务（异步，不等待）
        /// </summary>
        public void RequestCancel()
        {
            try
            {
                _disposeCts?.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // 忽略已释放的情况
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 只发送取消信号，不等待任务完成
                    // 避免阻塞请求响应
                    try
                    {
                        _disposeCts?.Cancel();
                    }
                    catch (ObjectDisposedException)
                    {
                        // 忽略
                    }

                    _disposeCts?.Dispose();
                }

                _disposed = true;
            }
        }

        #endregion
    }
}
