using System.Collections.Generic;
using System.Threading.Tasks;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务接口，支持多种 AI 提供商
    /// </summary>
    public interface ILLMService
    {
        /// <summary>
        /// 服务商名称
        /// </summary>
        string ProviderName { get; }

        /// <summary>
        /// 发送消息并获取响应
        /// </summary>
        Task<string> SendMessageAsync(string userMessage);

        /// <summary>
        /// 发送带历史记录的消息
        /// </summary>
        Task<string> SendMessagesWithHistoryAsync(List<object> messages);

        /// <summary>
        /// 发送纠错审阅请求（system + user）
        /// </summary>
        Task<string> SendProofreadMessageAsync(string systemContent, string userContent);
    }

    /// <summary>
    /// AI 提供商类型
    /// </summary>
    public enum AIProvider
    {
        DeepSeek,
        GLM,
        Ollama
    }
}
