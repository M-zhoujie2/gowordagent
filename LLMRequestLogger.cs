using System;
using System.IO;
using System.Linq;
using System.Text;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 请求日志信息
    /// 注意：出于安全考虑，此类不记录 API Key 和 HTTP Headers
    /// </summary>
    public class RequestLogInfo
    {
        public string Provider { get; set; }
        public DateTime RequestTime { get; set; }
        public DateTime ResponseTime { get; set; }
        public long ElapsedMs { get; set; }
        
        /// <summary>
        /// API Key 脱敏提示（前4位+***），用于调试识别
        /// </summary>
        public string ApiKeyHint { get; set; }
        
        public string SystemPrompt { get; set; }
        public string UserContent { get; set; }
        public int UserContentLength { get; set; }
        public string ResponseContent { get; set; }
        public int ResponseLength { get; set; }
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 设置 API Key 并生成脱敏提示
        /// </summary>
        public void SetApiKey(string apiKey)
        {
            if (string.IsNullOrEmpty(apiKey))
            {
                ApiKeyHint = "(空)";
            }
            else
            {
                int showLength = Math.Min(4, apiKey.Length);
                ApiKeyHint = apiKey.Substring(0, showLength) + new string('*', apiKey.Length - showLength);
            }
        }
    }

    /// <summary>
    /// LLM 请求日志记录器
    /// </summary>
    public class LLMRequestLogger
    {
        private readonly string _logFilePath;

        public LLMRequestLogger(string logFilePath = null)
        {
            _logFilePath = logFilePath ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GOWordAgent",
                "llm_requests.log");
            
            EnsureLogDirectoryExists();
        }

        /// <summary>
        /// 写入日志
        /// 注意：出于安全考虑，日志中不记录完整 API Key
        /// </summary>
        public void WriteLog(RequestLogInfo logInfo)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("========================================");
                sb.AppendLine($"时间: {logInfo.RequestTime:yyyy-MM-dd HH:mm:ss.fff}");
                sb.AppendLine($"提供商: {logInfo.Provider}");
                sb.AppendLine($"API Key: {logInfo.ApiKeyHint ?? "(空)"}");
                sb.AppendLine($"状态: {(logInfo.IsSuccess ? "成功" : "失败")}");
                sb.AppendLine($"耗时: {logInfo.ElapsedMs}ms");
                sb.AppendLine($"请求文本长度: {logInfo.UserContentLength} 字符");
                sb.AppendLine($"响应文本长度: {logInfo.ResponseLength} 字符");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【System Prompt】");
                sb.AppendLine(logInfo.SystemPrompt ?? "(空)");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【User Content】");
                sb.AppendLine(logInfo.UserContent ?? "(空)");
                sb.AppendLine("----------------------------------------");
                sb.AppendLine("【Response】");
                sb.AppendLine(logInfo.ResponseContent ?? "(空)");
                sb.AppendLine("========================================");
                sb.AppendLine();

                File.AppendAllText(_logFilePath, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LLMRequestLogger] 写入日志失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取日志文件路径
        /// </summary>
        public string GetLogFilePath()
        {
            return _logFilePath;
        }

        /// <summary>
        /// 获取最近的日志内容
        /// </summary>
        public string GetRecentLogs(int maxLines = 100)
        {
            try
            {
                if (!File.Exists(_logFilePath))
                    return "暂无日志";

                var lines = File.ReadAllLines(_logFilePath);
                if (lines.Length <= maxLines)
                    return string.Join("\n", lines);

                return string.Join("\n", lines.Skip(lines.Length - maxLines));
            }
            catch (Exception ex)
            {
                return $"读取日志失败: {ex.Message}";
            }
        }

        private void EnsureLogDirectoryExists()
        {
            try
            {
                var directory = Path.GetDirectoryName(_logFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LLMRequestLogger] 创建日志目录失败: {ex.Message}");
            }
        }
    }
}
