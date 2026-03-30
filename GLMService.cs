using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 智谱 GLM 服务
    /// </summary>
    public class GLMService : BaseLLMService
    {
        public override string ProviderName => "智谱 AI";

        /// <summary>
        /// GLM 校对需要更长的推理时间
        /// </summary>
        protected override int ProofreadTimeoutSeconds => 300; // 5分钟

        public GLMService(string apiKey, string apiUrl = null, string model = null)
            : base(apiKey, apiUrl, model,
                  "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                  "glm-4.7")
        {
        }

        /// <summary>
        /// GLM 特有参数：禁用思考模式以获得更快响应
        /// </summary>
        protected override Dictionary<string, object> BuildProofreadRequestBodyDict(List<object> messages)
        {
            var dict = base.BuildProofreadRequestBodyDict(messages);
            dict["enable_thinking"] = false; // GLM 特有参数
            return dict;
        }
    }
}
