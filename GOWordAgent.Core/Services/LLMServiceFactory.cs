using System;
using System.Collections.Generic;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务工厂
    /// </summary>
    public static class LLMServiceFactory
    {
        public static ILLMService CreateService(AIProvider provider, string apiKey, string apiUrl = null, string model = null)
        {
            switch (provider)
            {
                case AIProvider.DeepSeek:
                    return new DeepSeekService(apiKey, apiUrl, model);
                case AIProvider.GLM:
                    return new GLMService(apiKey, apiUrl, model);
                case AIProvider.Ollama:
                    return new OllamaService(apiUrl ?? "http://localhost:11434", apiKey);
                default:
                    throw new ArgumentException($"不支持的 AI 提供商: {provider}");
            }
        }

        public static Dictionary<AIProvider, string> GetProviders()
        {
            return new Dictionary<AIProvider, string>
            {
                { AIProvider.DeepSeek, "DeepSeek" },
                { AIProvider.GLM, "智谱 AI (GLM)" },
                { AIProvider.Ollama, "本地 Ollama" }
            };
        }

        public static string GetDefaultApiUrl(AIProvider provider)
        {
            switch (provider)
            {
                case AIProvider.DeepSeek:
                    return "https://api.deepseek.com/v1/chat/completions";
                case AIProvider.GLM:
                    return "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                case AIProvider.Ollama:
                    return "http://localhost:11434/api/chat";
                default:
                    return string.Empty;
            }
        }

        public static string GetDefaultModel(AIProvider provider)
        {
            switch (provider)
            {
                case AIProvider.DeepSeek:
                    return "deepseek-chat";
                case AIProvider.GLM:
                    return "glm-4.7";
                case AIProvider.Ollama:
                    return "llama2";
                default:
                    return "default";
            }
        }
    }
}
