using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务抽象基类
    /// </summary>
    public abstract class BaseLLMService : ILLMService, IDisposable
    {
        protected readonly string _apiUrl;
        protected readonly string _apiKey;
        protected readonly string _model;
        protected readonly HttpClient _httpClient;
        protected readonly LLMRequestLogger _logger;

        public abstract string ProviderName { get; }

        protected BaseLLMService(string apiKey, string apiUrl, string model, string defaultApiUrl, string defaultModel)
        {
            _apiKey = apiKey;
            _apiUrl = string.IsNullOrWhiteSpace(apiUrl) ? defaultApiUrl : apiUrl;
            _model = string.IsNullOrWhiteSpace(model) ? defaultModel : model;

            // 使用 HttpClientFactory 创建客户端，共享连接池
            _httpClient = HttpClientFactory.CreateAuthenticatedClient(apiKey, _apiUrl);

            _logger = new LLMRequestLogger();
        }

        public virtual async Task<string> SendMessageAsync(string userMessage, CancellationToken cancellationToken = default)
        {
            var messages = new List<object> { new { role = "user", content = userMessage } };
            return await SendRequestAsync(messages, cancellationToken);
        }

        public virtual async Task<string> SendMessagesWithHistoryAsync(List<object> messages, CancellationToken cancellationToken = default)
        {
            return await SendRequestAsync(messages, cancellationToken);
        }

        /// <summary>
        /// 校对请求超时时间（秒）- 校对场景可能需要更长时间
        /// </summary>
        protected virtual int ProofreadTimeoutSeconds => 300; // 默认5分钟

        public virtual async Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default)
        {
            var requestInfo = new RequestLogInfo
            {
                Provider = ProviderName,
                RequestTime = DateTime.Now,
                SystemPrompt = systemContent,
                UserContent = userContent,
                UserContentLength = userContent?.Length ?? 0
            };
            requestInfo.SetApiKey(_apiKey); // 记录脱敏的 API Key 提示

            var messages = new List<object>
            {
                new { role = "system", content = systemContent },
                new { role = "user", content = userContent }
            };

            string jsonContent = BuildProofreadRequestBody(messages);
            
            // 校对请求使用更长的超时时间
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                cts.CancelAfter(TimeSpan.FromSeconds(ProofreadTimeoutSeconds));
                string response = await PostAsync(jsonContent, requestInfo, cts.Token);
                return response;
            }
        }

        protected virtual async Task<string> SendRequestAsync(List<object> messages, CancellationToken cancellationToken = default)
        {
            string jsonContent = BuildRequestBody(messages);
            return await PostAsync(jsonContent, null, cancellationToken);
        }

        protected virtual string ParseResponse(JObject jsonResponse)
        {
            return jsonResponse["choices"]?[0]?["message"]?.Value<string>("content") ?? "未获取到回复内容";
        }

        /// <summary>
        /// 构建请求体 - 返回 Dictionary 以便子类灵活扩展
        /// </summary>
        protected virtual Dictionary<string, object> BuildRequestBodyDict(List<object> messages)
        {
            return new Dictionary<string, object>
            {
                ["model"] = _model,
                ["messages"] = messages,
                ["temperature"] = 0.7,
                ["max_tokens"] = 2000,
                ["stream"] = false
            };
        }

        /// <summary>
        /// 构建校对请求体 - 返回 Dictionary 以便子类灵活扩展
        /// </summary>
        protected virtual Dictionary<string, object> BuildProofreadRequestBodyDict(List<object> messages)
        {
            return new Dictionary<string, object>
            {
                ["model"] = _model,
                ["messages"] = messages,
                ["temperature"] = 0.1,
                ["max_tokens"] = 4000,
                ["stream"] = false
            };
        }

        protected virtual string BuildRequestBody(List<object> messages)
        {
            var dict = BuildRequestBodyDict(messages);
            return JsonConvert.SerializeObject(dict);
        }

        protected virtual string BuildProofreadRequestBody(List<object> messages)
        {
            var dict = BuildProofreadRequestBodyDict(messages);
            return JsonConvert.SerializeObject(dict);
        }

        protected virtual async Task<string> PostAsync(string jsonContent, RequestLogInfo logInfo, CancellationToken cancellationToken = default)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            string responseBody = null;

            try
            {
                System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Request URL: {_apiUrl}");
                System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Request Body: {jsonContent}");

                using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
                {
                    HttpResponseMessage response = await _httpClient.PostAsync(_apiUrl, content, cancellationToken)
                        .ConfigureAwait(false);
                    
                    responseBody = await response.Content.ReadAsStringAsync()
                        .ConfigureAwait(false);

                    stopwatch.Stop();

                    System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Response Status: {response.StatusCode}");
                    System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Response Body: {responseBody}");

                    if (response.IsSuccessStatusCode)
                    {
                        JObject jsonResponse = JObject.Parse(responseBody);
                        string result = ParseResponse(jsonResponse);
                        
                        // 记录成功的请求日志
                        if (logInfo != null)
                        {
                            logInfo.ResponseTime = DateTime.Now;
                            logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                            logInfo.ResponseContent = result;
                            logInfo.ResponseLength = result?.Length ?? 0;
                            logInfo.IsSuccess = true;
                            _logger.WriteLog(logInfo);
                        }
                        
                        return result;
                    }

                    // 记录失败的请求日志
                    if (logInfo != null)
                    {
                        logInfo.ResponseTime = DateTime.Now;
                        logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                        logInfo.ResponseContent = $"HTTP {(int)response.StatusCode}: {responseBody}";
                        logInfo.IsSuccess = false;
                        _logger.WriteLog(logInfo);
                    }

                    throw new LLMServiceException(
                        $"API 调用失败: {response.StatusCode}",
                        response.StatusCode,
                        responseBody,
                        ProviderName);
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                stopwatch.Stop();
                string errorMsg = "请求已取消";
                
                if (logInfo != null)
                {
                    logInfo.ResponseTime = DateTime.Now;
                    logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                    logInfo.ResponseContent = errorMsg;
                    logInfo.IsSuccess = false;
                    _logger.WriteLog(logInfo);
                }
                
                // 重新抛出，让上层处理取消
                throw;
            }
            catch (TaskCanceledException ex)
            {
                stopwatch.Stop();
                string errorMsg = "请求超时，请检查网络连接或稍后重试";
                
                if (logInfo != null)
                {
                    logInfo.ResponseTime = DateTime.Now;
                    logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                    logInfo.ResponseContent = errorMsg;
                    logInfo.IsSuccess = false;
                    _logger.WriteLog(logInfo);
                }
                
                throw new LLMServiceException(errorMsg, ex, ProviderName);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                string errorMsg = $"发生错误: {ex.Message}";
                
                if (logInfo != null)
                {
                    logInfo.ResponseTime = DateTime.Now;
                    logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                    logInfo.ResponseContent = errorMsg;
                    logInfo.IsSuccess = false;
                    _logger.WriteLog(logInfo);
                }
                
                throw new LLMServiceException(errorMsg, ex, ProviderName);
            }
        }

        #region 日志记录

        /// <summary>
        /// 获取日志文件路径
        /// </summary>
        public string GetLogFilePath() => _logger.GetLogFilePath();

        /// <summary>
        /// 获取最近的日志内容
        /// </summary>
        public string GetRecentLogs(int maxLines = 100) => _logger.GetRecentLogs(maxLines);

        #endregion

        #region IDisposable

        private bool _disposed = false;

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
                    // 注意：_httpClient 由 HttpClientFactory 管理，共享连接池
                    // 不应该在这里释放，避免影响其他实例
                }
                _disposed = true;
            }
        }

        /// <summary>
        /// 终结器
        /// </summary>
        ~BaseLLMService()
        {
            Dispose(false);
        }

        #endregion
    }

}
