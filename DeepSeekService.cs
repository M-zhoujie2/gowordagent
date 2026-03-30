using System.Threading;
using System.Threading.Tasks;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// DeepSeek API 服务
    /// </summary>
    public class DeepSeekService : BaseLLMService
    {
        public override string ProviderName => "DeepSeek";

        public DeepSeekService(string apiKey, string apiUrl = null, string model = null)
            : base(apiKey, apiUrl, model,
                  "https://api.deepseek.com/v1/chat/completions",
                  "deepseek-chat")
        {
        }

        public override async Task<string> SendProofreadMessageAsync(string systemContent, string userContent, CancellationToken cancellationToken = default)
        {
            // DeepSeek 特殊处理：使用思考模式
            return await base.SendProofreadMessageAsync(systemContent, userContent, cancellationToken);
        }
    }
}
