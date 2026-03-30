using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    /// 应用配置管理器（使用 DPAPI + HMAC 对本地配置文件进行加密和完整性保护）
    /// </summary>
    public static class ConfigManager
    {
        private static readonly string ConfigDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GOWordAgentAddIn");

        private static readonly string ConfigFile = Path.Combine(ConfigDir, "config.dat");
        
        // HMAC 密钥文件（用于完整性校验）
        private static readonly string HmacKeyFile = Path.Combine(ConfigDir, "config.key");

        /// <summary>
        /// 当前配置
        /// </summary>
        public static AIConfig CurrentConfig { get; private set; } = new AIConfig();

        /// <summary>
        /// 获取或生成 HMAC 密钥
        /// </summary>
        private static byte[] GetOrCreateHmacKey()
        {
            try
            {
                if (File.Exists(HmacKeyFile))
                {
                    string base64 = File.ReadAllText(HmacKeyFile, Encoding.UTF8);
                    return Convert.FromBase64String(base64);
                }
            }
            catch
            {
                // 密钥文件损坏，重新生成
            }

            // 生成新的随机密钥
            byte[] key = new byte[32]; // 256-bit key for HMAC-SHA256
            using (var rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(key);
            }

            try
            {
                Directory.CreateDirectory(ConfigDir);
                File.WriteAllText(HmacKeyFile, Convert.ToBase64String(key), Encoding.UTF8);
            }
            catch
            {
                // 如果无法保存密钥文件，使用派生密钥作为后备
                key = DeriveKeyFromMachineInfo();
            }

            return key;
        }

        /// <summary>
        /// 从机器信息派生密钥（后备方案）
        /// </summary>
        private static byte[] DeriveKeyFromMachineInfo()
        {
            // 使用机器名和用户SID作为派生基础
            string machineInfo = Environment.MachineName + Environment.UserDomainName;
            using (var sha256 = SHA256.Create())
            {
                return sha256.ComputeHash(Encoding.UTF8.GetBytes(machineInfo));
            }
        }

        /// <summary>
        /// 计算 HMAC-SHA256
        /// </summary>
        private static byte[] ComputeHmac(byte[] data, byte[] key)
        {
            using (var hmac = new HMACSHA256(key))
            {
                return hmac.ComputeHash(data);
            }
        }

        /// <summary>
        /// 加载配置（验证 HMAC 完整性后解密）
        /// </summary>
        public static void LoadConfig()
        {
            try
            {
                if (File.Exists(ConfigFile))
                {
                    string base64 = File.ReadAllText(ConfigFile, Encoding.UTF8);
                    byte[] combined = Convert.FromBase64String(base64);
                    
                    // 分离 HMAC 和加密数据
                    const int hmacLength = 32; // HMAC-SHA256 长度
                    if (combined.Length < hmacLength)
                    {
                        throw new ConfigSecurityException("配置文件损坏：数据长度不足");
                    }
                    
                    byte[] hmac = new byte[hmacLength];
                    byte[] encrypted = new byte[combined.Length - hmacLength];
                    Buffer.BlockCopy(combined, 0, hmac, 0, hmacLength);
                    Buffer.BlockCopy(combined, hmacLength, encrypted, 0, encrypted.Length);
                    
                    // 验证 HMAC
                    byte[] hmacKey = GetOrCreateHmacKey();
                    byte[] computedHmac = ComputeHmac(encrypted, hmacKey);
                    
                    if (!hmac.SequenceEqual(computedHmac))
                    {
                        throw new ConfigSecurityException(
                            "配置文件完整性校验失败。可能原因：\n" +
                            "1. 配置文件被篡改\n" +
                            "2. 配置文件损坏\n" +
                            "3. 配置在其他用户账户下创建\n\n" +
                            "建议：删除配置文件后重新配置。",
                            new Exception("HMAC 不匹配"));
                    }
                    
                    // HMAC 验证通过，解密数据
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
            catch (ConfigSecurityException)
            {
                // 安全异常，向上抛出以便 UI 层提示用户
                CurrentConfig = new AIConfig();
                throw;
            }
            catch (Exception ex)
            {
                // 其他异常（如 DPAPI 解密失败），使用默认配置
                CurrentConfig = new AIConfig();
                System.Diagnostics.Debug.WriteLine($"加载配置失败: {ex.Message}");
                
                // 如果是配置文件损坏，给出明确提示
                if (ex is CryptographicException)
                {
                    throw new ConfigSecurityException(
                        "配置文件解密失败。可能原因：\n" +
                        "1. 配置文件在其他 Windows 账户下创建\n" +
                        "2. 配置文件损坏\n\n" +
                        "建议：删除配置文件后重新配置。",
                        ex);
                }
            }
        }

        /// <summary>
        /// 保存配置（DPAPI 加密 + HMAC 完整性校验）
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
                
                // DPAPI 加密
                byte[] encrypted = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                
                // 计算 HMAC（对加密后的数据）
                byte[] hmacKey = GetOrCreateHmacKey();
                byte[] hmac = ComputeHmac(encrypted, hmacKey);
                
                // 组合：HMAC (32字节) + 加密数据
                byte[] combined = new byte[hmac.Length + encrypted.Length];
                Buffer.BlockCopy(hmac, 0, combined, 0, hmac.Length);
                Buffer.BlockCopy(encrypted, 0, combined, hmac.Length, encrypted.Length);
                
                // Base64 编码并保存
                string base64 = Convert.ToBase64String(combined);
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
