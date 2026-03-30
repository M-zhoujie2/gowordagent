using System.Collections.Generic;
using Newtonsoft.Json;

namespace GOWordAgentAddIn
{
    public class GLMService : BaseLLMService
    {
        public override string ProviderName => "智谱 AI";

        public GLMService(string apiKey, string apiUrl = null, string model = null)
            : base(apiKey, apiUrl, model,
                  "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                  "glm-4.7")
        {
        }

        protected override string BuildProofreadRequestBody(List<object> messages)
        {
            var requestBody = new
            {
                model = _model,
                enable_thinking = false,
                messages = messages,
                temperature = 0.1,
                max_tokens = 4000,
                stream = false
            };
            return JsonConvert.SerializeObject(requestBody);
        }
    }
}
