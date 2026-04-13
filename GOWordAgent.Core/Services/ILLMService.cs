using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务接口
    /// </summary>
    public interface ILLMService : IDisposable
    {
        /// <summary>
        /// 提供商名称
        /// </summary>
        string ProviderName { get; }

        /// <summary>
        /// 发送单条消息
        /// </summary>
        Task<string> SendMessageAsync(string userMessage, CancellationToken cancellationToken = default);

        /// <summary>
        /// 发送带历史记录的消息
        /// </summary>
        Task<string> SendMessagesWithHistoryAsync(List<object> messages, CancellationToken cancellationToken = default);

        /// <summary>
        /// 发送校对请求
        /// </summary>
        Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default);
    }
}
