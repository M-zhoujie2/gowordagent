using System.Collections.Generic;
using System.Threading;
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
        /// <param name="userMessage">用户消息</param>
        /// <param name="cancellationToken">取消令牌</param>
        Task<string> SendMessageAsync(string userMessage, CancellationToken cancellationToken = default);

        /// <summary>
        /// 发送带历史记录的消息
        /// </summary>
        /// <param name="messages">消息历史</param>
        /// <param name="cancellationToken">取消令牌</param>
        Task<string> SendMessagesWithHistoryAsync(List<object> messages, CancellationToken cancellationToken = default);

        /// <summary>
        /// 发送纠错审阅请求（system + user）
        /// </summary>
        /// <param name="systemContent">系统提示词</param>
        /// <param name="userContent">用户内容</param>
        /// <param name="cancellationToken">取消令牌</param>
        Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default);
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
