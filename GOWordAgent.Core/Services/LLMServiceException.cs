using System;
using System.Net;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务异常
    /// </summary>
    public class LLMServiceException : Exception
    {
        /// <summary>
        /// HTTP 状态码
        /// </summary>
        public HttpStatusCode? StatusCode { get; }

        /// <summary>
        /// 原始响应内容
        /// </summary>
        public string? ResponseBody { get; }

        /// <summary>
        /// 提供商名称
        /// </summary>
        public string ProviderName { get; }

        public LLMServiceException(string message, string providerName) 
            : base(message)
        {
            ProviderName = providerName;
        }

        public LLMServiceException(string message, Exception inner, string providerName) 
            : base(message, inner)
        {
            ProviderName = providerName;
        }

        public LLMServiceException(string message, HttpStatusCode statusCode, string responseBody, string providerName) 
            : base(message)
        {
            StatusCode = statusCode;
            ResponseBody = responseBody;
            ProviderName = providerName;
        }

        /// <summary>
        /// 获取友好的错误信息
        /// </summary>
        public string GetFriendlyErrorMessage()
        {
            if (StatusCode.HasValue)
            {
                switch (StatusCode.Value)
                {
                    case HttpStatusCode.Unauthorized:
                        return $"[{ProviderName}] API Key 无效或已过期，请检查配置";
                    case HttpStatusCode.TooManyRequests:
                        return $"[{ProviderName}] 请求过于频繁，请稍后再试";
                    case HttpStatusCode.InternalServerError:
                    case HttpStatusCode.BadGateway:
                    case HttpStatusCode.ServiceUnavailable:
                        return $"[{ProviderName}] 服务暂时不可用，请稍后再试";
                }
            }

            if (Message.Contains("timeout", StringComparison.OrdinalIgnoreCase))
            {
                return $"[{ProviderName}] 请求超时，请检查网络连接";
            }

            return $"[{ProviderName}] {Message}";
        }
    }
}
