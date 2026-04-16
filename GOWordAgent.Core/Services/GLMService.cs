namespace GOWordAgentAddIn
{
    /// <summary>
    /// 智谱 GLM API 服务
    /// </summary>
    public class GLMService : BaseLLMService
    {
        public override string ProviderName => "智谱 AI";

        public GLMService(string apiKey, string? apiUrl = null, string? model = null)
            : base(apiKey, apiUrl, model,
                  "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                  "glm-4.7")
        {
        }
    }
}
