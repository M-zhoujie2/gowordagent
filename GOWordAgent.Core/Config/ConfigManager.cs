using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

    public class ProviderConfig
    {
        public string ApiKey { get; set; } = "";
        public string ApiUrl { get; set; } = "";
        public string Model { get; set; } = "";
    }

    public class AIConfig
    {
        public AIProvider Provider { get; set; } = AIProvider.DeepSeek;
        public string ApiKey { get; set; } = "";
        public string ApiUrl { get; set; } = "";
        public string Model { get; set; } = "";
        public bool AutoConnect { get; set; } = false;
        public Dictionary<string, ProviderConfig> ProviderConfigs { get; set; } = new Dictionary<string, ProviderConfig>();
        public string ProofreadMode { get; set; } = "精准校验";
        public string ProofreadPrecisePrompt { get; set; } = "";
        public string ProofreadFullTextPrompt { get; set; } = "";
        public string ProofreadPrompt { get; set; } = "";
        public string PrivacyConsentLastShownDate { get; set; } = "";
    }

    /// <summary>
    /// 跨平台配置管理器
    /// </summary>
    public static class ConfigManager
    {
        private static string ConfigDir => RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GOWordAgentAddIn")
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config", "gowordagent");

        private static string ConfigFile => Path.Combine(ConfigDir, "config.dat");

        public static AIConfig CurrentConfig { get; private set; } = new AIConfig();

        public static void LoadConfig()
        {
            try
            {
                if (!File.Exists(ConfigFile))
                {
                    CurrentConfig = new AIConfig();
                    return;
                }

                byte[] encrypted = File.ReadAllBytes(ConfigFile);
                byte[] decrypted;

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                }
                else
                {
                    decrypted = LinuxCrypto.Decrypt(encrypted);
                }

                string json = Encoding.UTF8.GetString(decrypted);
                var config = JsonConvert.DeserializeObject<AIConfig>(json);

                if (config != null)
                {
                    if (config.ProviderConfigs == null)
                        config.ProviderConfigs = new Dictionary<string, ProviderConfig>();

                    CurrentConfig = config;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载配置失败: {ex.Message}");
                CurrentConfig = new AIConfig();
            }
        }

        public static void SaveConfig(AIConfig config)
        {
            try
            {
                if (config.ProviderConfigs == null)
                    config.ProviderConfigs = new Dictionary<string, ProviderConfig>();

                Directory.CreateDirectory(ConfigDir);

                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                byte[] encrypted;

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                }
                else
                {
                    encrypted = LinuxCrypto.Encrypt(plainBytes);
                }

                File.WriteAllBytes(ConfigFile, encrypted);
                CurrentConfig = config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存配置失败: {ex.Message}");
                throw;
            }
        }

        public static void SaveCurrentConfig(AIProvider provider, string apiKey, string apiUrl, string model, bool autoConnect = false)
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

        public static ProviderConfig? GetProviderConfig(AIProvider provider)
        {
            if (CurrentConfig?.ProviderConfigs == null)
                return null;

            CurrentConfig.ProviderConfigs.TryGetValue(provider.ToString(), out var cfg);
            return cfg;
        }

        public static string GetConfigFilePath() => ConfigFile;

        public static (string mode, string prompt) GetProofreadConfig()
        {
            var config = CurrentConfig;
            string mode = string.IsNullOrEmpty(config.ProofreadMode) ? "精准校验" : config.ProofreadMode;
            string prompt = mode == "全文校验" ? config.ProofreadFullTextPrompt : config.ProofreadPrecisePrompt;

            return (mode, prompt);
        }

        public static string GetProofreadPromptForMode(string mode)
        {
            var config = CurrentConfig;
            string saved = mode == "全文校验" ? config.ProofreadFullTextPrompt : config.ProofreadPrecisePrompt;
            if (!string.IsNullOrEmpty(saved)) return saved;
            return mode == "全文校验" ? DefaultProofreadPrompts.FullText : DefaultProofreadPrompts.Precise;
        }
    }

    /// <summary>
    /// Linux 加密实现 (AES-GCM + /etc/machine-id)
    /// </summary>
    public static class LinuxCrypto
    {
        private static byte[] DeriveKey()
        {
            var machineId = File.ReadAllText("/etc/machine-id").Trim();
            return SHA256.HashData(Encoding.UTF8.GetBytes(machineId));
        }

        public static byte[] Encrypt(byte[] plainData)
        {
            var key = DeriveKey();
            var nonce = RandomNumberGenerator.GetBytes(12);

            using var aes = new AesGcm(key, 16);
            var cipherData = new byte[plainData.Length];
            var tag = new byte[16];

            aes.Encrypt(nonce, plainData, cipherData, tag);

            var result = new byte[12 + 16 + cipherData.Length];
            Buffer.BlockCopy(nonce, 0, result, 0, 12);
            Buffer.BlockCopy(tag, 0, result, 12, 16);
            Buffer.BlockCopy(cipherData, 0, result, 28, cipherData.Length);
            return result;
        }

        public static byte[] Decrypt(byte[] encryptedData)
        {
            var key = DeriveKey();
            var nonce = new byte[12];
            var tag = new byte[16];
            var cipherData = new byte[encryptedData.Length - 28];

            Buffer.BlockCopy(encryptedData, 0, nonce, 0, 12);
            Buffer.BlockCopy(encryptedData, 12, tag, 0, 16);
            Buffer.BlockCopy(encryptedData, 28, cipherData, 0, cipherData.Length);

            using var aes = new AesGcm(key, 16);
            var plainData = new byte[cipherData.Length];
            aes.Decrypt(nonce, cipherData, tag, plainData);
            return plainData;
        }
    }
}
