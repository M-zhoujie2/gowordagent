using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using GOWordAgentAddIn.Models;

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

    }

    /// <summary>
    /// 跨平台配置管理器
    /// </summary>
    public static class ConfigManager
    {
        public static string ConfigDir => RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GOWordAgentAddIn")
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config", "gowordagent");

        private static string ConfigFile => Path.Combine(ConfigDir, "config.dat");

        public static AIConfig CurrentConfig { get; private set; } = new AIConfig();

        public static void LoadConfig()
        {
            try
            {
                Directory.CreateDirectory(ConfigDir);
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

                // 检查磁盘空间（至少 1MB）
                try
                {
                    var configDirInfo = new DirectoryInfo(ConfigDir);
                    var driveInfo = new DriveInfo(configDirInfo.Root.FullName);
                    if (driveInfo.AvailableFreeSpace < 1024 * 1024)
                    {
                        throw new IOException("磁盘空间不足，无法保存配置");
                    }
                }
                catch (Exception ex) when (ex is not IOException)
                {
                    // 在某些文件系统上 DriveInfo 可能不可用，忽略非 IO 异常
                }

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

                var tempFile = ConfigFile + ".tmp";
                File.WriteAllBytes(tempFile, encrypted);
                File.Move(tempFile, ConfigFile, overwrite: true);
                SetUnixFilePermissions(ConfigFile, 600);
                CurrentConfig = config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存配置失败: {ex.Message}");
                throw;
            }
        }

        private static void SetUnixFilePermissions(string path, int mode)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && !RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                return;
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo("chmod")
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false
                };
                psi.ArgumentList.Add(Convert.ToString(mode, 8));
                psi.ArgumentList.Add(path);
                using var proc = System.Diagnostics.Process.Start(psi);
                proc?.WaitForExit(2000);
            }
            catch { /* 忽略权限设置失败 */ }
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

        public static string GetProofreadPromptForMode(string mode)
        {
            var config = CurrentConfig;
            string saved = mode == "全文校验" ? config.ProofreadFullTextPrompt : config.ProofreadPrecisePrompt;
            if (!string.IsNullOrEmpty(saved)) return saved;
            return mode == "全文校验" ? DefaultProofreadPrompts.FullText : DefaultProofreadPrompts.Precise;
        }
    }

    /// <summary>
    /// Linux 加密实现 (AES-GCM + /etc/machine-id + 随机 salt)
    /// </summary>
    public static class LinuxCrypto
    {
        private static readonly string SaltFile = Path.Combine(ConfigManager.ConfigDir, ".salt");

        /// <summary>
        /// 从机器 ID 和 salt 派生密钥
        /// </summary>
        private static byte[] DeriveKey()
        {
            string? machineId = TryGetMachineId();
            byte[] salt = EnsureSalt();

            if (string.IsNullOrWhiteSpace(machineId))
            {
                var homeDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                machineId = homeDir + Environment.UserName;
            }

            // 使用 HKDF-like 组合：先分别哈希，再组合哈希
            var machineBytes = Encoding.UTF8.GetBytes(machineId);
            using var hmac = new HMACSHA256(salt);
            return hmac.ComputeHash(machineBytes);
        }

        /// <summary>
        /// 确保 salt 文件存在（首次运行时生成 32 字节随机 salt）
        /// </summary>
        private static byte[] EnsureSalt()
        {
            try
            {
                Directory.CreateDirectory(ConfigManager.ConfigDir);
                if (File.Exists(SaltFile))
                {
                    var salt = File.ReadAllBytes(SaltFile);
                    if (salt.Length >= 16) return salt;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LinuxCrypto] 读取 salt 文件失败: {ex.Message}");
            }

            var newSalt = RandomNumberGenerator.GetBytes(32);
            try
            {
                var temp = SaltFile + ".tmp";
                File.WriteAllBytes(temp, newSalt);
                File.Move(temp, SaltFile, overwrite: true);
                // 设置仅所有者可读写
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) || RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo("chmod")
                        {
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false
                        };
                        psi.ArgumentList.Add("600");
                        psi.ArgumentList.Add(SaltFile);
                        using var proc = System.Diagnostics.Process.Start(psi);
                        proc?.WaitForExit(2000);
                    }
                    catch (Exception chmodEx)
                    {
                        Debug.WriteLine($"[LinuxCrypto] 设置 salt 文件权限失败: {chmodEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LinuxCrypto] 写入 salt 文件失败: {ex.Message}");
                // 如果旧 salt 不存在且写入失败，后续将无法解密配置，应抛出异常
                if (!File.Exists(SaltFile))
                {
                    throw new InvalidOperationException($"无法创建加密 salt 文件 '{SaltFile}'，配置将无法安全保存。请检查磁盘空间和权限。", ex);
                }
            }
            return newSalt;
        }

        /// <summary>
        /// 尝试获取机器 ID（尝试多种来源）
        /// </summary>
        private static string? TryGetMachineId()
        {
            string[] possiblePaths = new[]
            {
                "/etc/machine-id",
                "/var/lib/dbus/machine-id",
                "/sys/class/dmi/id/product_uuid",
                "/proc/sys/kernel/random/boot_id"
            };

            foreach (var path in possiblePaths)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        var content = File.ReadAllText(path).Trim();
                        if (!string.IsNullOrWhiteSpace(content) && content.Length >= 16)
                        {
                            return content;
                        }
                    }
                }
                catch { /* 忽略单个文件的错误，继续尝试下一个 */ }
            }

            try
            {
                var hostName = System.Net.Dns.GetHostName();
                if (!string.IsNullOrWhiteSpace(hostName))
                {
                    return hostName;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 加密数据
        /// </summary>
        public static byte[] Encrypt(byte[] plainData)
        {
            if (plainData == null)
                throw new ArgumentNullException(nameof(plainData));

            if (plainData.Length == 0)
                return Array.Empty<byte>();

            try
            {
                var key = DeriveKey();
                var nonce = RandomNumberGenerator.GetBytes(12);

                using var aes = new AesGcm(key, 16);
                var cipherData = new byte[plainData.Length];
                var tag = new byte[16];

                aes.Encrypt(nonce, plainData, cipherData, tag);

                // 组合: nonce(12) + tag(16) + ciphertext
                var result = new byte[12 + 16 + cipherData.Length];
                Buffer.BlockCopy(nonce, 0, result, 0, 12);
                Buffer.BlockCopy(tag, 0, result, 12, 16);
                Buffer.BlockCopy(cipherData, 0, result, 28, cipherData.Length);
                return result;
            }
            catch (CryptographicException ex)
            {
                throw new InvalidOperationException("Encryption failed", ex);
            }
        }

        /// <summary>
        /// 解密数据
        /// </summary>
        public static byte[] Decrypt(byte[] encryptedData)
        {
            if (encryptedData == null)
                throw new ArgumentNullException(nameof(encryptedData));

            if (encryptedData.Length == 0)
                return Array.Empty<byte>();

            if (encryptedData.Length < 28)
                throw new ArgumentException(
                    "Invalid encrypted data format. Data too short.",
                    nameof(encryptedData));

            try
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
            catch (CryptographicException ex)
            {
                throw new InvalidOperationException(
                    "Decryption failed. The configuration file may be corrupted or was created on a different machine.",
                    ex);
            }
        }

    }
}
