using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对缓存管理器（基于 ConcurrentDictionary，减少锁竞争；内存统计为 O(1) 原子计数）
    /// </summary>
    public static class ProofreadCacheManager
    {
        private static readonly ConcurrentDictionary<string, ParagraphResult> _globalCache = new ConcurrentDictionary<string, ParagraphResult>();
        private static readonly ConcurrentDictionary<string, long> _accessCount = new ConcurrentDictionary<string, long>();
        private static long _accessCounter = 0;
        private static long _estimatedBytes = 0;

        private const int MaxCacheSize = 1000;
        private const long MaxCacheBytes = 50 * 1024 * 1024; // 50MB 内存上限

        public static event EventHandler<CacheHitEventArgs>? OnCacheHit;

        public static bool TryGetCachedResult(string text, int index, out ParagraphResult? result, string? mode = null)
        {
            result = null;
            if (string.IsNullOrEmpty(text)) return false;

            var cacheKey = ComputeHash(text, mode);

            if (_globalCache.TryGetValue(cacheKey, out var cached))
            {
                Debug.WriteLine($"[ProofreadCacheManager] 段落 {index + 1} 命中缓存!");
                _accessCount[cacheKey] = Interlocked.Increment(ref _accessCounter);

                result = new ParagraphResult
                {
                    Index = index,
                    OriginalText = text,
                    ResultText = cached.ResultText,
                    IsCompleted = true,
                    IsCached = true,
                    ProcessTime = DateTime.UtcNow,
                    ElapsedMs = 0,
                    Items = cached.Items?.ToList() ?? new List<ProofreadIssueItem>()
                };

                OnCacheHit?.Invoke(null, new CacheHitEventArgs { Index = index, CacheKey = cacheKey });
                return true;
            }

            Debug.WriteLine($"[ProofreadCacheManager] 缓存状态: 共 {_globalCache.Count} 项");
            return false;
        }

        public static void StoreResult(string text, ParagraphResult result, string? mode = null)
        {
            if (string.IsNullOrEmpty(text) || result == null) return;

            var cacheKey = ComputeHash(text, mode);
            var itemBytes = EstimateResultBytes(result);

            // 检查内存使用，如果超过上限则淘汰旧条目
            while (Interlocked.Read(ref _estimatedBytes) + itemBytes > MaxCacheBytes && !_globalCache.IsEmpty)
            {
                EvictOldestEntries(10);
            }

            if (_globalCache.Count >= MaxCacheSize && !_globalCache.ContainsKey(cacheKey))
            {
                EvictOldestEntries(100);
            }

            if (_globalCache.TryAdd(cacheKey, result))
            {
                _accessCount[cacheKey] = Interlocked.Increment(ref _accessCounter);
                Interlocked.Add(ref _estimatedBytes, itemBytes);
                Debug.WriteLine($"[ProofreadCacheManager] 已存入缓存, 缓存总数={_globalCache.Count}, 估计内存: {Interlocked.Read(ref _estimatedBytes) / 1024}KB");
            }
        }

        private static void EvictOldestEntries(int count)
        {
            var keysToRemove = _accessCount
                .OrderBy(kvp => kvp.Value)
                .Take(count)
                .Select(kvp => kvp.Key)
                .ToList();

            long freedBytes = 0;
            foreach (var key in keysToRemove)
            {
                if (_globalCache.TryRemove(key, out var removed))
                {
                    freedBytes += EstimateResultBytes(removed);
                }
                _accessCount.TryRemove(key, out _);
            }
            Interlocked.Add(ref _estimatedBytes, -freedBytes);
            Debug.WriteLine($"[ProofreadCacheManager] 已淘汰 {keysToRemove.Count} 个旧缓存条目，释放 {freedBytes / 1024}KB");
        }

        public static string ComputeHash(string text, string? mode = null)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            var content = string.IsNullOrEmpty(mode) ? text : $"[{mode}]{text}";
            var bytes = Encoding.UTF8.GetBytes(content);

            using (var sha256 = SHA256.Create())
            {
                var hash = sha256.ComputeHash(bytes);
                return Convert.ToBase64String(hash);
            }
        }

        public static (int Count, long EstimatedBytes) GetCacheStats()
        {
            return (_globalCache.Count, Interlocked.Read(ref _estimatedBytes));
        }

        /// <summary>
        /// 更精确的内存估算：包含字符串 UTF-16 字节、列表引用开销、对象头
        /// </summary>
        private static long EstimateResultBytes(ParagraphResult? r)
        {
            if (r == null) return 0;

            long bytes = 0;
            // 字符串 UTF-16 长度 * 2
            bytes += (r.OriginalText?.Length ?? 0) * 2L;
            bytes += (r.ResultText?.Length ?? 0) * 2L;

            // ParagraphResult 对象头 + 字段引用（64位约 56 字节，32位约 32 字节）
            bytes += IntPtr.Size == 8 ? 56 : 32;

            // Items 列表开销
            var items = r.Items;
            if (items != null)
            {
                bytes += IntPtr.Size == 8 ? 32 : 20; // List<T> 对象头
                bytes += items.Count * (IntPtr.Size == 8 ? 8 : 4); // 引用数组
                foreach (var item in items)
                {
                    bytes += IntPtr.Size == 8 ? 64 : 40; // ProofreadIssueItem 对象头 + 字段
                    bytes += (item.Type?.Length ?? 0) * 2L;
                    bytes += (item.Original?.Length ?? 0) * 2L;
                    bytes += (item.Modified?.Length ?? 0) * 2L;
                    bytes += (item.Reason?.Length ?? 0) * 2L;
                    bytes += (item.Severity?.Length ?? 0) * 2L;
                }
            }

            return bytes;
        }

        public static void ClearCache()
        {
            _globalCache.Clear();
            _accessCount.Clear();
            Interlocked.Exchange(ref _accessCounter, 0);
            Interlocked.Exchange(ref _estimatedBytes, 0);
            Debug.WriteLine("[ProofreadCacheManager] 缓存已清除");
        }
    }

    public class CacheHitEventArgs : EventArgs
    {
        public int Index { get; set; }
        public string CacheKey { get; set; } = "";
    }
}
