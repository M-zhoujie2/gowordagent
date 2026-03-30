using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GOWordAgentAddIn
{
    public class OllamaService : ILLMService
    {
        private readonly string _apiUrl;
        private readonly string _model;
        private readonly HttpClient _httpClient;
        private readonly string _logFilePath;

        public string ProviderName => "Ollama(本地)";

        public OllamaService(string apiUrl, string model = null)
        {
            _apiUrl = apiUrl ?? "http://localhost:11434";
            _model = model ?? "llama2";

            var handler = new HttpClientHandler
            {
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                UseProxy = false,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(120) };
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            // 初始化日志文件路径
            _logFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GOWordAgent",
                "llm_requests.log");
            
            EnsureLogDirectoryExists();
        }

        public async Task<string> SendMessageAsync(string userMessage)
        {
            var messages = new List<object> { new { role = "user", content = userMessage } };
            return await SendRequestAsync(messages);
        }

        public async Task<string> SendMessagesWithHistoryAsync(List<object> messages)
        {
            return await SendRequestAsync(messages);
        }

        public async Task<string> SendProofreadMessageAsync(string systemContent, string userContent)
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

            return await SendRequestAsync(messages, temperature: 0.1, logInfo: requestInfo);
        }

        private async Task<string> SendRequestAsync(List<object> messages, double temperature = 0.7, RequestLogInfo logInfo = null)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                var requestBody = new
                {
                    model = _model,
                    messages = messages,
                    stream = false,
                    options = new { temperature, num_predict = 4000 }
                };

                string jsonContent = JsonConvert.SerializeObject(requestBody);
                System.Diagnostics.Debug.WriteLine($"[Ollama] Request URL: {_apiUrl}/api/chat");
                System.Diagnostics.Debug.WriteLine($"[Ollama] Request Body: {jsonContent}");

                using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
                {
                    var response = await _httpClient.PostAsync($"{_apiUrl.TrimEnd('/')}/api/chat", content).ConfigureAwait(false);
                    string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    stopwatch.Stop();

                    System.Diagnostics.Debug.WriteLine($"[Ollama] Response Status: {response.StatusCode}");
                    System.Diagnostics.Debug.WriteLine($"[Ollama] Response Body: {responseBody}");

                    if (response.IsSuccessStatusCode)
                    {
                        var jsonResponse = JObject.Parse(responseBody);
                        string result = jsonResponse["message"]?.Value<string>("content") ?? "未获取到回复内容";

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
                string errorMsg = "请求超时，请确认 Ollama 是否已启动";

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

        #endregion
    }

}
