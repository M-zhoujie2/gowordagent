namespace GOWordAgentAddIn
{
    /// <summary>
    /// DeepSeek API 服务
    /// </summary>
    public class DeepSeekService : BaseLLMService
    {
        public override string ProviderName => "DeepSeek";

        public DeepSeekService(string apiKey, string? apiUrl = null, string? model = null)
            : base(apiKey, apiUrl, model,
                  "https://api.deepseek.com/v1/chat/completions",
                  "deepseek-chat")
        {
        }
    }
}
