using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// Ollama 本地 API 服务
    /// </summary>
    public class OllamaService : BaseLLMService
    {
        public override string ProviderName => "Ollama(本地)";

        /// <summary>
        /// 创建 Ollama 服务实例
        /// </summary>
        /// <param name="apiUrl">Ollama 服务地址，如 http://localhost:11434</param>
        /// <param name="model">模型名称，如 llama2</param>
        public OllamaService(string apiUrl, string model = null)
            : base(apiKey: null, apiUrl: apiUrl, model: model,
                  defaultApiUrl: "http://localhost:11434/api/chat",
                  defaultModel: "llama2")
        {
            // Ollama 不需要 API Key，移除默认的 Authorization 头
            _httpClient.DefaultRequestHeaders.Remove("Authorization");
        }

        /// <summary>
        /// 构建请求体（Ollama 格式）
        /// </summary>
        protected override string BuildRequestBody(List<object> messages)
        {
            var requestBody = new
            {
                model = _model,
                messages = messages,
                stream = false,
                options = new { temperature = 0.7, num_predict = 2000 }
            };
            return JsonConvert.SerializeObject(requestBody);
        }

        /// <summary>
        /// 构建校对请求体（Ollama 格式，更低温度）
        /// </summary>
        protected override string BuildProofreadRequestBody(List<object> messages)
        {
            var requestBody = new
            {
                model = _model,
                messages = messages,
                stream = false,
                options = new { temperature = 0.1, num_predict = 4000 }
            };
            return JsonConvert.SerializeObject(requestBody);
        }

        /// <summary>
        /// 解析响应（Ollama 格式）
        /// </summary>
        protected override string ParseResponse(JObject jsonResponse)
        {
            // Ollama 响应格式: { "message": { "role": "assistant", "content": "..." } }
            return jsonResponse["message"]?.Value<string>("content") ?? "未获取到回复内容";
        }

        /// <summary>
        /// 发送消息（添加 Ollama 特定的超时处理）
        /// </summary>
        public override async Task<string> SendMessageAsync(string userMessage, CancellationToken cancellationToken = default)
        {
            try
            {
                return await base.SendMessageAsync(userMessage, cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException ex)
            {
                throw new LLMServiceException("请求超时，请确认 Ollama 是否已启动", ex, ProviderName);
            }
            catch (HttpRequestException ex)
            {
                throw new LLMServiceException($"连接失败，请确认 Ollama 服务是否运行在 {_apiUrl}", ex, ProviderName);
            }
        }

        /// <summary>
        /// 发送校对请求（添加 Ollama 特定的超时处理）
        /// </summary>
        public override async Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default)
        {
            try
            {
                return await base.SendProofreadMessageAsync(systemContent, userContent, cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException ex)
            {
                throw new LLMServiceException("请求超时，请确认 Ollama 是否已启动", ex, ProviderName);
            }
            catch (HttpRequestException ex)
            {
                throw new LLMServiceException($"连接失败，请确认 Ollama 服务是否运行在 {_apiUrl}", ex, ProviderName);
            }
        }
    }
}
