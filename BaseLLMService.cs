using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
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
        private readonly string _logFilePath;

        public abstract string ProviderName { get; }

        protected BaseLLMService(string apiKey, string apiUrl, string model, string defaultApiUrl, string defaultModel)
        {
            _apiKey = apiKey;
            _apiUrl = string.IsNullOrWhiteSpace(apiUrl) ? defaultApiUrl : apiUrl;
            _model = string.IsNullOrWhiteSpace(model) ? defaultModel : model;

            var handler = new HttpClientHandler
            {
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 |
                               System.Security.Authentication.SslProtocols.Tls13,
                UseProxy = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                MaxConnectionsPerServer = 10
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
            
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            _httpClient.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate");
            
            if (!string.IsNullOrEmpty(apiKey))
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            }

            // 初始化日志文件路径
            _logFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GOWordAgent",
                "llm_requests.log");
            
            EnsureLogDirectoryExists();
        }

        public virtual async Task<string> SendMessageAsync(string userMessage)
        {
            var messages = new List<object> { new { role = "user", content = userMessage } };
            return await SendRequestAsync(messages);
        }

        public virtual async Task<string> SendMessagesWithHistoryAsync(List<object> messages)
        {
            return await SendRequestAsync(messages);
        }

        public virtual async Task<string> SendProofreadMessageAsync(string systemContent, string userContent)
        {
            var requestInfo = new RequestLogInfo
            {
                Provider = ProviderName,
                RequestTime = DateTime.Now,
                SystemPrompt = systemContent,
                UserContent = userContent,
                UserContentLength = userContent?.Length ?? 0
            };

            var messages = new List<object>
            {
                new { role = "system", content = systemContent },
                new { role = "user", content = userContent }
            };

            string jsonContent = BuildProofreadRequestBody(messages);
            string response = await PostAsync(jsonContent, requestInfo);
            
            return response;
        }

        protected virtual async Task<string> SendRequestAsync(List<object> messages)
        {
            string jsonContent = BuildRequestBody(messages);
            return await PostAsync(jsonContent, null);
        }

        protected virtual string ParseResponse(JObject jsonResponse)
        {
            return jsonResponse["choices"]?[0]?["message"]?.Value<string>("content") ?? "未获取到回复内容";
        }

        protected virtual string BuildRequestBody(List<object> messages)
        {
            var requestBody = new
            {
                model = _model,
                messages = messages,
                temperature = 0.7,
                max_tokens = 2000,
                stream = false
            };
            return JsonConvert.SerializeObject(requestBody);
        }

        protected virtual string BuildProofreadRequestBody(List<object> messages)
        {
            var requestBody = new
            {
                model = _model,
                messages = messages,
                temperature = 0.1,
                max_tokens = 4000,
                stream = false
            };
            return JsonConvert.SerializeObject(requestBody);
        }

        protected virtual async Task<string> PostAsync(string jsonContent, RequestLogInfo logInfo)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            string responseBody = null;

            try
            {
                System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Request URL: {_apiUrl}");
                System.Diagnostics.Debug.WriteLine($"[{ProviderName}] Request Body: {jsonContent}");

                using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
                {
                    HttpResponseMessage response = await _httpClient.PostAsync(_apiUrl, content)
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
                            WriteLog(logInfo);
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
                        WriteLog(logInfo);
                    }

                    return $"API 调用失败: {response.StatusCode}\n{responseBody}";
                }
            }
            catch (TaskCanceledException)
            {
                stopwatch.Stop();
                string errorMsg = "请求超时，请检查网络连接或稍后重试";
                
                if (logInfo != null)
                {
                    logInfo.ResponseTime = DateTime.Now;
                    logInfo.ElapsedMs = stopwatch.ElapsedMilliseconds;
                    logInfo.ResponseContent = errorMsg;
                    logInfo.IsSuccess = false;
                    WriteLog(logInfo);
                }
                
                return errorMsg;
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
                    WriteLog(logInfo);
                }
                
                return errorMsg;
            }
        }

        #region 日志记录

        private void EnsureLogDirectoryExists()
        {
            try
            {
                var directory = Path.GetDirectoryName(_logFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch { }
        }

        private void WriteLog(RequestLogInfo logInfo)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("========================================");
                sb.AppendLine($"时间: {logInfo.RequestTime:yyyy-MM-dd HH:mm:ss.fff}");
                sb.AppendLine($"提供商: {logInfo.Provider}");
                sb.AppendLine($"状态: {(logInfo.IsSuccess ? "成功" : "失败")}");
                sb.AppendLine($"耗时: {logInfo.ElapsedMs}ms");
                sb.AppendLine($"请求文本长度: {logInfo.UserContentLength} 字符");
                sb.AppendLine($"响应文本长度: {logInfo.ResponseLength} 字符");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【System Prompt】");
                sb.AppendLine(logInfo.SystemPrompt ?? "(空)");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【User Content】");
                sb.AppendLine(logInfo.UserContent ?? "(空)");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【Response】");
                sb.AppendLine(logInfo.ResponseContent ?? "(空)");
                sb.AppendLine("========================================");
                sb.AppendLine();

                File.AppendAllText(_logFilePath, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Log Error] 写入日志失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取日志文件路径
        /// </summary>
        public string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// 获取最近的日志内容
        /// </summary>
        public string GetRecentLogs(int maxLines = 100)
        {
            try
            {
                if (!File.Exists(_logFilePath))
                    return "暂无日志";

                var lines = File.ReadAllLines(_logFilePath);
                if (lines.Length <= maxLines)
                    return string.Join("\n", lines);

                return string.Join("\n", lines.Skip(lines.Length - maxLines));
            }
            catch (Exception ex)
            {
                return $"读取日志失败: {ex.Message}";
            }
        }

        #endregion
    }

    /// <summary>
    /// 请求日志信息
    /// </summary>
    public class RequestLogInfo
    {
        public string Provider { get; set; }
        public DateTime RequestTime { get; set; }
        public DateTime ResponseTime { get; set; }
        public long ElapsedMs { get; set; }
        public string SystemPrompt { get; set; }
        public string UserContent { get; set; }
        public int UserContentLength { get; set; }
        public string ResponseContent { get; set; }
        public int ResponseLength { get; set; }
        public bool IsSuccess { get; set; }
    }
}
