using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 默认提示词常量
    /// </summary>
    public static class DefaultProofreadPrompts
    {
        public const string Precise = "你是中文校对编辑。精确纠错，仅校对文字错误（错别字、多字、漏字、语病、术语不一致），不校对标点、换行、空格等。\n\n" +
            "规则：只标错误不改风格，每处给原文/修改/理由，不确定标[待确认]。不改原文观点/风格/术语体系，无错误输出\"未发现错误\"。\n\n" +
            "格式：\n【第X处】类型：...｜严重度：high/medium/low\n原文：...\n修改：...\n理由：...";

        public const string FullText = "你是中文校对编辑。校对以下文本中的：\n" +
            "- 错别字、多字、漏字\n" +
            "- 语法错误、句式杂糅、搭配不当\n" +
            "- 标点错误（中英文混用、多余、缺失）\n" +
            "- 序号不连续、不统一\n" +
            "- 用词不当、口语化表述\n" +
            "- 专业术语前后不一致\n\n" +
            "规则：只标错误不改风格，每处给原文/修改/理由，不确定标[待确认]。不改原文观点/风格/术语体系，无错误输出\"未发现错误\"。\n\n" +
            "格式：\n【第X处】类型：...\n原文：...\n修改：...\n理由：...";
    }
    /// <summary>
    /// 单个服务商的配置参数
    /// </summary>
    public class ProviderConfig
    {
        public string ApiKey { get; set; } = "";
        public string ApiUrl { get; set; } = "";
        public string Model { get; set; } = "";
    }

    /// <summary>
    /// AI 配置数据类
    /// </summary>
    public class AIConfig
    {
        /// <summary>
        /// 当前选中的 AI 提供商类型
        /// </summary>
        public AIProvider Provider { get; set; } = AIProvider.DeepSeek;

        /// <summary>
        /// 当前 API Key（兼容旧配置）
        /// </summary>
        public string ApiKey { get; set; } = "";

        /// <summary>
        /// 当前 API URL（兼容旧配置）
        /// </summary>
        public string ApiUrl { get; set; } = "";

        /// <summary>
        /// 当前模型名称（兼容旧配置）
        /// </summary>
        public string Model { get; set; } = "";

        /// <summary>
        /// 是否自动连接
        /// </summary>
        public bool AutoConnect { get; set; } = false;

        /// <summary>
        /// 按服务商保存的配置字典
        /// </summary>
        public Dictionary<string, ProviderConfig> ProviderConfigs { get; set; } = new Dictionary<string, ProviderConfig>();

        /// <summary>
        /// 校验模式（精准校验/全文校验）
        /// </summary>
        public string ProofreadMode { get; set; } = "精准校验";

        /// <summary>
        /// 精准校验提示词
        /// </summary>
        public string ProofreadPrecisePrompt { get; set; } = "";

        /// <summary>
        /// 全文校验提示词
        /// </summary>
        public string ProofreadFullTextPrompt { get; set; } = "";

        /// <summary>
        /// 校验提示词（兼容旧配置）
        /// </summary>
        public string ProofreadPrompt { get; set; } = "";
    }

    /// <summary>
    /// 应用配置管理器（使用 DPAPI 对本地配置文件进行加解密）
    /// </summary>
    public static class ConfigManager
    {
        private static readonly string ConfigDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GOWordAgentAddIn");

        private static readonly string ConfigFile = Path.Combine(ConfigDir, "config.dat");

        /// <summary>
        /// 当前配置
        /// </summary>
        public static AIConfig CurrentConfig { get; private set; } = new AIConfig();

        /// <summary>
        /// 加载配置
        /// </summary>
        public static void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    string base64 = File.ReadAllText(ConfigFile, Encoding.UTF8);
                    byte[] encrypted = Convert.FromBase64String(base64);
                    byte[] decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                    string json = Encoding.UTF8.GetString(decrypted);
                    var config = JsonConvert.DeserializeObject<AIConfig>(json);
                    if (config != null)
                    {
                        // 向后兼容：旧配置没有 ProviderConfigs，将顶层配置迁移进去
                        if (config.ProviderConfigs == null)
                            config.ProviderConfigs = new Dictionary<string, ProviderConfig>();

                        if (config.ProviderConfigs.Count == 0 && !string.IsNullOrWhiteSpace(config.ApiKey))
                        {
                            config.ProviderConfigs[config.Provider.ToString()] = new ProviderConfig
                            {
                                ApiKey = config.ApiKey,
                                ApiUrl = config.ApiUrl,
                                Model = config.Model
                            };
                        }

                        CurrentConfig = config;
                    }
                }
            }
            catch (Exception ex)
            {
                // 加载失败时使用默认配置
                CurrentConfig = new AIConfig();
                System.Diagnostics.Debug.WriteLine($"加载配置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public static void SaveConfig(AIConfig config)
        {
            try
            {
                if (config.ProviderConfigs == null)
                    config.ProviderConfigs = new Dictionary<string, ProviderConfig>();

                // 确保目录存在
                Directory.CreateDirectory(ConfigDir);

                // 序列化为 JSON
                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                byte[] encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                string base64 = Convert.ToBase64String(encrypted);
                File.WriteAllText(ConfigFile, base64, Encoding.UTF8);

                // 更新当前配置
                CurrentConfig = config;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存配置失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 保存当前配置（从 UI 控件读取），同时更新对应服务商的字典条目
        /// </summary>
        public static void SaveCurrentConfig(
            AIProvider provider,
            string apiKey,
            string apiUrl,
            string model,
            bool autoConnect = false)
        {
            var config = CurrentConfig;
            config.Provider = provider;
            config.ApiKey = apiKey ?? "";
            config.ApiUrl = apiUrl ?? "";
            config.Model = model ?? "";
            config.AutoConnect = autoConnect;

            if (config.ProviderConfigs == null)
                config.ProviderConfigs = new Dictionary<string, ProviderConfig>();

            config.ProviderConfigs[provider.ToString()] = new ProviderConfig
            {
                ApiKey = apiKey ?? "",
                ApiUrl = apiUrl ?? "",
                Model = model ?? ""
            };

            SaveConfig(config);
        }

        /// <summary>
        /// 获取指定服务商的保存配置，若不存在返回 null
        /// </summary>
        public static ProviderConfig GetProviderConfig(AIProvider provider)
        {
            if (CurrentConfig?.ProviderConfigs == null)
                return null;

            CurrentConfig.ProviderConfigs.TryGetValue(provider.ToString(), out var cfg);
            return cfg;
        }

        /// <summary>
        /// 获取配置文件路径
        /// </summary>
        public static string GetConfigFilePath()
        {
            return ConfigFile;
        }

        /// <summary>
        /// 保存校验配置
        /// </summary>
        public static void SaveProofreadConfig(string mode, string prompt)
        {
            var config = CurrentConfig;
            config.ProofreadMode = mode;
            // 根据模式保存到对应的字段
            if (mode == "全文校验")
                config.ProofreadFullTextPrompt = prompt;
            else
                config.ProofreadPrecisePrompt = prompt;
            SaveConfig(config);
        }

        /// <summary>
        /// 获取校验配置
        /// </summary>
        public static (string mode, string prompt) GetProofreadConfig()
        {
            var config = CurrentConfig;
            string mode = string.IsNullOrEmpty(config.ProofreadMode) ? "精准校验" : config.ProofreadMode;
            string prompt = mode == "全文校验" ? config.ProofreadFullTextPrompt : config.ProofreadPrecisePrompt;
            
            return (mode, prompt);
        }

        /// <summary>
        /// 获取指定模式的提示词（含默认值）
        /// </summary>
        public static string GetProofreadPromptForMode(string mode)
        {
            var config = CurrentConfig;
            string saved = mode == "全文校验" ? config.ProofreadFullTextPrompt : config.ProofreadPrecisePrompt;
            if (!string.IsNullOrEmpty(saved)) return saved;
            return mode == "全文校验" ? DefaultProofreadPrompts.FullText : DefaultProofreadPrompts.Precise;
        }
    }
}
