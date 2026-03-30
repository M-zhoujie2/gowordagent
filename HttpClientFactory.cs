using System;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// HttpClient 工厂 - 管理共享的 HttpClient 实例
    /// 避免多次实例化导致的端口耗尽和 DNS 变更问题
    /// </summary>
    public static class HttpClientFactory
    {
        // 静态共享的 HttpClient 实例（按基础地址缓存）
        private static readonly ConcurrentDictionary<string, HttpClient> _clients = 
            new ConcurrentDictionary<string, HttpClient>();
        
        // 共享的 Handler 配置
        private static readonly HttpClientHandler _sharedHandler = new HttpClientHandler
        {
            SslProtocols = System.Security.Authentication.SslProtocols.Tls12 |
                           System.Security.Authentication.SslProtocols.Tls13,
            UseProxy = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
            MaxConnectionsPerServer = 20  // 增加连接池大小
        };

        /// <summary>
        /// 获取或创建 HttpClient 实例
        /// </summary>
        /// <param name="baseAddress">基础地址（用于缓存键）</param>
        /// <returns>共享的 HttpClient 实例</returns>
        public static HttpClient GetClient(string baseAddress)
        {
            if (string.IsNullOrWhiteSpace(baseAddress))
                throw new ArgumentException("基础地址不能为空", nameof(baseAddress));

            // 规范化基础地址（去掉路径，只保留 scheme://host:port）
            var normalizedKey = NormalizeBaseAddress(baseAddress);

            return _clients.GetOrAdd(normalizedKey, key =>
            {
                var client = new HttpClient(_sharedHandler, disposeHandler: false)
                {
                    Timeout = TimeSpan.FromSeconds(120),
                    BaseAddress = new Uri(key)
                };
                
                // 设置默认请求头
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate");
                
                return client;
            });
        }

        /// <summary>
        /// 创建带有认证头的 HttpClient（每个服务实例独立，但共享 Handler）
        /// </summary>
        /// <param name="apiKey">API 密钥</param>
        /// <param name="apiUrl">完整 API URL</param>
        /// <param name="timeoutSeconds">超时时间（秒），默认120秒</param>
        /// <returns>配置好的 HttpClient</returns>
        public static HttpClient CreateAuthenticatedClient(string apiKey, string apiUrl, int timeoutSeconds = 120)
        {
            // 使用共享 Handler 创建新的 HttpClient 实例
            // 这样每个服务可以有自己的认证头，但底层连接是共享的
            var client = new HttpClient(_sharedHandler, disposeHandler: false)
            {
                Timeout = TimeSpan.FromSeconds(timeoutSeconds)
            };
            
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate");
            
            if (!string.IsNullOrEmpty(apiKey))
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            }
            
            return client;
        }

        /// <summary>
        /// 规范化基础地址（提取 scheme://host:port）
        /// </summary>
        private static string NormalizeBaseAddress(string url)
        {
            try
            {
                var uri = new Uri(url);
                // 对于 API 端点，我们只缓存到主机级别
                // 例如：https://api.deepseek.com/v1/chat/completions -> https://api.deepseek.com/
                return $"{uri.Scheme}://{uri.Authority}/";
            }
            catch
            {
                return url.ToLowerInvariant();
            }
        }

        /// <summary>
        /// 清理所有缓存的 HttpClient（仅在需要时调用，如 DNS 变更）
        /// </summary>
        public static void ClearCache()
        {
            foreach (var client in _clients.Values)
            {
                client?.Dispose();
            }
            _clients.Clear();
        }

        /// <summary>
        /// 获取当前缓存的客户端数量（调试用）
        /// </summary>
        public static int CachedClientCount => _clients.Count;
    }
}
