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

        protected BaseLLMService(string apiKey, string? apiUrl, string? model, string defaultApiUrl, string defaultModel)
        {
            _apiKey = apiKey;
            _apiUrl = string.IsNullOrWhiteSpace(apiUrl) ? defaultApiUrl : apiUrl;
            _model = string.IsNullOrWhiteSpace(model) ? defaultModel : model;

            // 使用共享 HttpClient 工厂，避免 Socket 耗尽（Linux 下 TIME_WAIT 堆积问题）
            _httpClient = SharedHttpClientFactory.GetOrCreate(ProviderName, _apiUrl, _apiKey);
            // 如果 client 已存在但 API Key 变更，更新 Authorization
            SharedHttpClientFactory.UpdateAuthorization(ProviderName, _apiUrl, _apiKey);
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
        protected virtual int ProofreadTimeoutSeconds => 180;

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

        /// <summary>
        /// 安全解析 JSON，返回 null 表示解析失败
        /// </summary>
        private static JObject? TryParseJson(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;
            try
            {
                var token = JToken.Parse(text);
                return token as JObject;
            }
            catch
            {
                return null;
            }
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
            const int maxRetries = 2;
            int attempt = 0;

            while (true)
            {
                attempt++;
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
                            JObject? jsonResponse = TryParseJson(responseBody);
                            if (jsonResponse != null)
                            {
                                return ParseResponse(jsonResponse);
                            }
                            throw new LLMServiceException(
                                "API 返回了非 JSON 内容，可能是网关错误或认证页拦截",
                                response.StatusCode,
                                responseBody,
                                ProviderName);
                        }

                        // 对可重试状态码进行指数退避重试
                        if (attempt <= maxRetries && IsRetryableStatusCode(response.StatusCode))
                        {
                            int delayMs = (int)Math.Pow(2, attempt) * 500;
                            await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
                            continue;
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
                catch (TaskCanceledException) when (attempt <= maxRetries)
                {
                    int delayMs = (int)Math.Pow(2, attempt) * 500;
                    await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
                    continue;
                }
                catch (LLMServiceException) when (attempt <= maxRetries)
                {
                    // 仅对超时类错误继续重试，业务错误直接抛出
                    throw;
                }
                catch (HttpRequestException) when (attempt <= maxRetries)
                {
                    int delayMs = (int)Math.Pow(2, attempt) * 500;
                    await Task.Delay(delayMs, cancellationToken).ConfigureAwait(false);
                    continue;
                }
                catch (Exception ex)
                {
                    throw new LLMServiceException($"发生错误: {ex.Message}", ex, ProviderName);
                }
            }
        }

        private static bool IsRetryableStatusCode(HttpStatusCode statusCode)
        {
            return statusCode == HttpStatusCode.BadGateway
                || statusCode == HttpStatusCode.ServiceUnavailable
                || statusCode == HttpStatusCode.GatewayTimeout
                || statusCode == HttpStatusCode.TooManyRequests
                || (int)statusCode == 499; // Client Closed Request (nginx)
        }

        #region IDisposable

        public void Dispose()
        {
            // HttpClient 由 SharedHttpClientFactory 全局复用，不在此处释放
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // HttpClient 由 SharedHttpClientFactory 全局复用，不在此处释放
        }

        #endregion
    }
}
