using System;
using System.Collections.Generic;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// LLM 服务工厂，用于创建不同的 AI 服务实例
    /// </summary>
    public static class LLMServiceFactory
    {
        /// <summary>
        /// 创建 LLM 服务实例
        /// </summary>
        public static ILLMService CreateService(AIProvider provider, string apiKey, string apiUrl = null, string model = null)
        {
            switch (provider)
            {
                case AIProvider.DeepSeek:
                    return new DeepSeekService(apiKey, apiUrl, model);
                case AIProvider.GLM:
                    return new GLMService(apiKey, apiUrl, model);
                case AIProvider.Ollama:
                    // Ollama: apiKey 是模型名称，apiUrl 是服务地址
                    return new OllamaService(apiUrl ?? "http://localhost:11434", apiKey);
                default:
                    throw new ArgumentException($"不支持的 AI 提供商: {provider}");
            }
        }

        /// <summary>
        /// 获取所有支持的 AI 提供商
        /// </summary>
        public static Dictionary<AIProvider, string> GetProviders()
        {
            return new Dictionary<AIProvider, string>
            {
                { AIProvider.DeepSeek, "DeepSeek" },
                { AIProvider.GLM, "智谱 AI (GLM)" },
                { AIProvider.Ollama, "本地 Ollama" }
            };
        }

        /// <summary>
        /// 获取默认 API URL
        /// </summary>
        public static string GetDefaultApiUrl(AIProvider provider)
        {
            switch (provider)
            {
                case AIProvider.DeepSeek:
                    return "https://api.deepseek.com/v1/chat/completions";
                case AIProvider.GLM:
                    return "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                case AIProvider.Ollama:
                    return "http://localhost:11434/api/generate";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// 获取默认模型名称
        /// </summary>
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
