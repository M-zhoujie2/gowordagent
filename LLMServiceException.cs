using System;
using System.Net;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务异常 - API 调用失败时抛出
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
        public string ResponseBody { get; }

        /// <summary>
        /// 提供商名称
        /// </summary>
        public string ProviderName { get; }

        public LLMServiceException(string message) : base(message)
        {
        }

        public LLMServiceException(string message, Exception innerException) 
            : base(message, innerException)
        {
        }

        public LLMServiceException(string message, Exception innerException, string providerName)
            : base(message, innerException)
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

        public LLMServiceException(string message, Exception innerException, HttpStatusCode statusCode, string responseBody, string providerName)
            : base(message, innerException)
        {
            StatusCode = statusCode;
            ResponseBody = responseBody;
            ProviderName = providerName;
        }

        /// <summary>
        /// 获取友好错误信息
        /// </summary>
        public string GetFriendlyErrorMessage()
        {
            if (StatusCode.HasValue)
            {
                switch (StatusCode.Value)
                {
                    case HttpStatusCode.Unauthorized:
                        return $"[{ProviderName}] API Key 无效或已过期，请检查配置";
                    case (HttpStatusCode)429: // TooManyRequests
                        return $"[{ProviderName}] 请求过于频繁，请稍后重试";
                    case HttpStatusCode.BadRequest:
                        return $"[{ProviderName}] 请求参数错误: {Message}";
                    case HttpStatusCode.InternalServerError:
                    case HttpStatusCode.BadGateway:
                    case HttpStatusCode.ServiceUnavailable:
                        return $"[{ProviderName}] 服务暂时不可用，请稍后重试";
                    default:
                        return $"[{ProviderName}] API 调用失败: {StatusCode} - {Message}";
                }
            }
            return $"[{ProviderName}] {Message}";
        }
    }
}
