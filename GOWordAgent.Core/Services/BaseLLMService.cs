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
    public abstract class BaseLLMService : ILLMService
    {
        protected readonly string _apiUrl;
        protected readonly string _apiKey;
        protected readonly string _model;
        protected readonly HttpClient _httpClient;

        public abstract string ProviderName { get; }

        protected BaseLLMService(string apiKey, string apiUrl, string model, string defaultApiUrl, string defaultModel)
        {
            _apiKey = apiKey;
            _apiUrl = string.IsNullOrWhiteSpace(apiUrl) ? defaultApiUrl : apiUrl;
            _model = string.IsNullOrWhiteSpace(model) ? defaultModel : model;

            // 创建 HttpClient
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
            
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            
            if (!string.IsNullOrEmpty(apiKey))
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            }
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
        /// 校对请求超时时间（秒）
        /// </summary>
        protected virtual int ProofreadTimeoutSeconds => 300;

        public virtual async Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default)
        {
            var messages = new List<object>
            {
                new { role = "system", content = systemContent },
                new { role = "user", content = userContent }
            };

            string jsonContent = BuildProofreadRequestBody(messages);
            
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
            {
                cts.CancelAfter(TimeSpan.FromSeconds(ProofreadTimeoutSeconds));
                return await PostAsync(jsonContent, cts.Token);
            }
        }

        protected virtual async Task<string> SendRequestAsync(List<object> messages, CancellationToken cancellationToken = default)
        {
            string jsonContent = BuildRequestBody(messages);
            return await PostAsync(jsonContent, cancellationToken);
        }

        protected virtual string ParseResponse(JObject jsonResponse)
        {
            return jsonResponse["choices"]?[0]?["message"]?.Value<string>("content") ?? "未获取到回复内容";
        }

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

        protected virtual async Task<string> PostAsync(string jsonContent, CancellationToken cancellationToken = default)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
                {
                    HttpResponseMessage response = await _httpClient.PostAsync(_apiUrl, content, cancellationToken)
                        .ConfigureAwait(false);
                    
                    string responseBody = await response.Content.ReadAsStringAsync()
                        .ConfigureAwait(false);

                    stopwatch.Stop();

                    if (response.IsSuccessStatusCode)
                    {
                        JObject jsonResponse = JObject.Parse(responseBody);
                        return ParseResponse(jsonResponse);
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
                throw;
            }
            catch (TaskCanceledException ex)
            {
                throw new LLMServiceException("请求超时，请检查网络连接或稍后重试", ex, ProviderName);
            }
            catch (Exception ex)
            {
                throw new LLMServiceException($"发生错误: {ex.Message}", ex, ProviderName);
            }
        }

        #region IDisposable

        private bool _disposed = false;

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
                    _httpClient?.Dispose();
                }
                _disposed = true;
            }
        }

        #endregion
    }
}
