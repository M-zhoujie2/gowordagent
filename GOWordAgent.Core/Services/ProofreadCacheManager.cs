using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对缓存管理器
    /// </summary>
    public static class ProofreadCacheManager
    {
        private static readonly Dictionary<string, ParagraphResult> _globalCache = new Dictionary<string, ParagraphResult>();
        private static readonly object _cacheLock = new object();
        private const int MaxCacheSize = 1000;
        private static readonly Dictionary<string, long> _accessCount = new Dictionary<string, long>();
        private static long _accessCounter = 0;

        public static event EventHandler<CacheHitEventArgs>? OnCacheHit;

        public static bool TryGetCachedResult(string text, int index, out ParagraphResult? result, string? mode = null)
        {
            result = null;
            if (string.IsNullOrEmpty(text)) return false;

            var cacheKey = ComputeHash(text, mode);
            ParagraphResult? cached = null;

            lock (_cacheLock)
            {
                Debug.WriteLine($"[ProofreadCacheManager] 缓存状态: 共 {_globalCache.Count} 项");
                _globalCache.TryGetValue(cacheKey, out cached);
            }

            if (cached != null)
            {
                Debug.WriteLine($"[ProofreadCacheManager] 段落 {index + 1} 命中缓存!");

                lock (_cacheLock)
                {
                    _accessCount[cacheKey] = System.Threading.Interlocked.Increment(ref _accessCounter);
                }

                result = new ParagraphResult
                {
                    Index = index,
                    OriginalText = text,
                    ResultText = cached.ResultText,
                    IsCompleted = true,
                    IsCached = true,
                    ProcessTime = DateTime.Now,
                    ElapsedMs = 0,
                    Items = cached.Items?.ToList() ?? new List<ProofreadIssueItem>()
                };

                OnCacheHit?.Invoke(null, new CacheHitEventArgs { Index = index, CacheKey = cacheKey });
                return true;
            }

            return false;
        }

        public static void StoreResult(string text, ParagraphResult result, string? mode = null)
        {
            if (string.IsNullOrEmpty(text) || result == null) return;

            var cacheKey = ComputeHash(text, mode);

            lock (_cacheLock)
            {
                if (_globalCache.Count >= MaxCacheSize && !_globalCache.ContainsKey(cacheKey))
                {
                    EvictOldestEntries(100);
                }

                if (!_globalCache.ContainsKey(cacheKey))
                {
                    _globalCache[cacheKey] = result;
                    _accessCount[cacheKey] = System.Threading.Interlocked.Increment(ref _accessCounter);
                    Debug.WriteLine($"[ProofreadCacheManager] 已存入缓存, 缓存总数={_globalCache.Count}");
                }
            }
        }

        private static void EvictOldestEntries(int count)
        {
            var keysToRemove = _accessCount.OrderBy(kvp => kvp.Value).Take(count).Select(kvp => kvp.Key).ToList();
            foreach (var key in keysToRemove)
            {
                _globalCache.Remove(key);
                _accessCount.Remove(key);
            }
            Debug.WriteLine($"[ProofreadCacheManager] 已淘汰 {keysToRemove.Count} 个旧缓存条目");
        }

        public static string ComputeHash(string text, string? mode = null)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            var content = string.IsNullOrEmpty(mode) ? text : $"[{mode}]{text}";
            var bytes = Encoding.UTF8.GetBytes(content);

            // 使用实例方法，避免线程安全问题
            using (var sha256 = SHA256.Create())
            {
                var hash = sha256.ComputeHash(bytes);
                return Convert.ToBase64String(hash);
            }
        }

        public static (int Count, long EstimatedBytes) GetCacheStats()
        {
            lock (_cacheLock)
            {
                int count = _globalCache.Count;
                long bytes = _globalCache.Values.Sum(r =>
                    (r.OriginalText?.Length ?? 0) + (r.ResultText?.Length ?? 0)) * 2;
                return (count, bytes);
            }
        }

        public static void ClearCache()
        {
            lock (_cacheLock)
            {
                _globalCache.Clear();
                _accessCount.Clear();
                System.Threading.Interlocked.Exchange(ref _accessCounter, 0);
                Debug.WriteLine("[ProofreadCacheManager] 缓存已清除");
            }
        }
    }

    public class CacheHitEventArgs : EventArgs
    {
        public int Index { get; set; }
        public string CacheKey { get; set; } = "";
    }
}
