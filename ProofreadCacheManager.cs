using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对缓存管理器 - 管理段落级缓存
    /// </summary>
    public static class ProofreadCacheManager
    {
        // 静态缓存，所有实例共享（跨校对会话保留）
        private static readonly Dictionary<string, ParagraphResult> _globalCache = new Dictionary<string, ParagraphResult>();
        private static readonly object _cacheLock = new object();
        
        // 最大缓存条目数，防止内存无限增长
        private const int MaxCacheSize = 1000;
        
        // 访问计数，用于LRU淘汰
        private static readonly Dictionary<string, long> _accessCount = new Dictionary<string, long>();
        private static long _accessCounter = 0;

        /// <summary>
        /// 缓存命中事件
        /// </summary>
        public static event EventHandler<CacheHitEventArgs> OnCacheHit;

        /// <summary>
        /// 尝试从缓存获取结果
        /// </summary>
        public static bool TryGetCachedResult(string text, int index, out ParagraphResult result)
        {
            result = null;
            if (string.IsNullOrEmpty(text)) return false;

            var cacheKey = ComputeHash(text);
            ParagraphResult cached = null;

            lock (_cacheLock)
            {
                Debug.WriteLine($"[ProofreadCacheManager] 缓存状态: 共 {_globalCache.Count} 项");
                _globalCache.TryGetValue(cacheKey, out cached);
            }

            if (cached != null)
            {
                Debug.WriteLine($"[ProofreadCacheManager] ✅ 段落 {index + 1} 命中缓存! Key={cacheKey.Substring(0, Math.Min(8, cacheKey.Length))}...");
                
                // 更新访问计数
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
                    Items = cached.Items?.ToList() ?? new List<ProofreadItem>()
                };

                OnCacheHit?.Invoke(null, new CacheHitEventArgs { Index = index, CacheKey = cacheKey });
                return true;
            }
            else
            {
                Debug.WriteLine($"[ProofreadCacheManager] ❌ 段落 {index + 1} 未命中缓存, Key={cacheKey.Substring(0, Math.Min(8, cacheKey.Length))}...");
                return false;
            }
        }

        /// <summary>
        /// 存储结果到缓存（带大小限制）
        /// </summary>
        public static void StoreResult(string text, ParagraphResult result)
        {
            if (string.IsNullOrEmpty(text) || result == null) return;

            var cacheKey = ComputeHash(text);

            lock (_cacheLock)
            {
                // 检查是否需要淘汰旧缓存
                if (_globalCache.Count >= MaxCacheSize && !_globalCache.ContainsKey(cacheKey))
                {
                    EvictOldestEntries(100); // 淘汰100个最旧的
                }
                
                if (!_globalCache.ContainsKey(cacheKey))
                {
                    _globalCache[cacheKey] = result;
                    _accessCount[cacheKey] = System.Threading.Interlocked.Increment(ref _accessCounter);
                    Debug.WriteLine($"[ProofreadCacheManager] 💾 已存入缓存, Key={cacheKey.Substring(0, Math.Min(8, cacheKey.Length))}..., 缓存总数={_globalCache.Count}");
                }
                else
                {
                    // 更新访问计数
                    _accessCount[cacheKey] = System.Threading.Interlocked.Increment(ref _accessCounter);
                    Debug.WriteLine($"[ProofreadCacheManager] ⚠️ 缓存已存在, 跳过存储");
                }
            }
        }
        
        /// <summary>
        /// 淘汰最旧的缓存条目
        /// </summary>
        private static void EvictOldestEntries(int count)
        {
            var keysToRemove = _accessCount.OrderBy(kvp => kvp.Value).Take(count).Select(kvp => kvp.Key).ToList();
            foreach (var key in keysToRemove)
            {
                _globalCache.Remove(key);
                _accessCount.Remove(key);
            }
            Debug.WriteLine($"[ProofreadCacheManager] 🗑️ 已淘汰 {keysToRemove.Count} 个旧缓存条目");
        }

        /// <summary>
        /// 计算文本哈希（用于缓存键）
        /// </summary>
        public static string ComputeHash(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(text);
                var hash = sha.ComputeHash(bytes);
                return Convert.ToBase64String(hash);
            }
        }

        /// <summary>
        /// 获取缓存统计
        /// </summary>
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

        /// <summary>
        /// 清除缓存
        /// </summary>
        public static void ClearCache()
        {
            lock (_cacheLock)
            {
                _globalCache.Clear();
                _accessCount.Clear();
                _accessCounter = 0;
                Debug.WriteLine("[ProofreadCacheManager] 缓存已清除");
            }
        }
    }

    /// <summary>
    /// 缓存命中事件参数
    /// </summary>
    public class CacheHitEventArgs : EventArgs
    {
        public int Index { get; set; }
        public string CacheKey { get; set; }
    }
}
