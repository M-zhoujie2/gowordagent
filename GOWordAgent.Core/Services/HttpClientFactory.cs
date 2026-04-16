using System;
using System.Collections.Concurrent;
using System.Net.Http;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// HttpClient 工厂 - 按 provider 复用连接池，避免 Socket 耗尽
    /// </summary>
    public static class SharedHttpClientFactory
    {
        private static readonly ConcurrentDictionary<string, HttpClient> _clients = new ConcurrentDictionary<string, HttpClient>();
        private static readonly SocketsHttpHandler _sharedHandler;

        static SharedHttpClientFactory()
        {
            _sharedHandler = new SocketsHttpHandler
            {
                MaxConnectionsPerServer = 20,
                UseProxy = false,
                PooledConnectionLifetime = TimeSpan.FromMinutes(5),
                EnableMultipleHttp2Connections = true
            };
        }

        /// <summary>
        /// 获取或创建 HttpClient（不随服务实例 Dispose）
        /// </summary>
        public static HttpClient GetOrCreate(string providerName, string apiUrl, string? apiKey)
        {
            var key = $"{providerName}:{apiUrl}";
            return _clients.GetOrAdd(key, _ =>
            {
                var client = new HttpClient(_sharedHandler, disposeHandler: false)
                {
                    Timeout = TimeSpan.FromSeconds(300)
                };
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                if (!string.IsNullOrEmpty(apiKey))
                {
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                }
                return client;
            });
        }

        /// <summary>
        /// 更新现有 client 的 Authorization（用于 API Key 变更时）
        /// </summary>
        public static void UpdateAuthorization(string providerName, string apiUrl, string? apiKey)
        {
            var key = $"{providerName}:{apiUrl}";
            if (_clients.TryGetValue(key, out var client))
            {
                client.DefaultRequestHeaders.Remove("Authorization");
                if (!string.IsNullOrEmpty(apiKey))
                {
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                }
            }
        }
    }
}
