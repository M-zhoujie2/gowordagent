using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
using System.Diagnostics;
using GOWordAgentAddIn;
using GOWordAgentAddIn.Models;

namespace GOWordAgent.WpsService.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProofreadController : ControllerBase
    {
        private readonly ILogger<ProofreadController> _logger;

        public ProofreadController(ILogger<ProofreadController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        [EnableRateLimiting("ProofreadPolicy")]
        public async Task<IActionResult> Proofread([FromBody] ProofreadRequest request)
        {
            // 每次新请求前重置全局取消状态，避免上次取消永久阻塞后续请求
            ProofreadService.ResetGlobalCancellation();

            try
            {
                // 输入验证
                if (request == null)
                {
                    return BadRequest(new { error = "请求体不能为空" });
                }

                if (request.Paragraphs == null || request.Paragraphs.Count == 0)
                {
                    return BadRequest(new { error = "文档段落不能为空" });
                }

                // 检查段落数据有效性
                for (int i = 0; i < request.Paragraphs.Count; i++)
                {
                    var para = request.Paragraphs[i];
                    if (para.StartOffset < 0 || para.EndOffset < para.StartOffset)
                    {
                        return BadRequest(new { error = $"第 {i+1} 段偏移量无效" });
                    }
                }

                _logger.LogInformation("Starting proofread for document with {ParagraphCount} paragraphs", 
                    request.Paragraphs.Count);

                // 获取有效的 API Key
                var apiKey = !string.IsNullOrWhiteSpace(request.ApiKey) 
                    ? request.ApiKey 
                    : ConfigManager.CurrentConfig.ApiKey;

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    return BadRequest(new { error = "API Key 未配置" });
                }

                // 获取有效的提供商配置
                var provider = request.Provider;
                var apiUrl = !string.IsNullOrWhiteSpace(request.ApiUrl) 
                    ? request.ApiUrl 
                    : ConfigManager.CurrentConfig.ApiUrl;
                var model = !string.IsNullOrWhiteSpace(request.Model) 
                    ? request.Model 
                    : ConfigManager.CurrentConfig.Model;

                // 创建临时的 LLM 服务（使用请求中的配置）
                if (!Enum.TryParse<AIProvider>(provider, out var providerEnum))
                {
                    return BadRequest(new { error = $"无效的 AI 提供商: {provider}" });
                }
                using var llmService = LLMServiceFactory.CreateService(providerEnum, apiKey, apiUrl, model);

                // 获取提示词
                var prompt = string.IsNullOrEmpty(request.Prompt) 
                    ? ConfigManager.GetProofreadPromptForMode(request.Mode.ToString())
                    : request.Prompt;

                // 创建校对服务
                using var proofreadService = new ProofreadService(
                    llmService,
                    prompt,
                    concurrency: 5,
                    proofreadMode: request.Mode.ToString()
                );

                // 合并段落文本
                var fullText = string.Join("\n", request.Paragraphs.Select(p => p.Text));

                if (string.IsNullOrWhiteSpace(fullText))
                {
                    return BadRequest(new { error = "文档内容为空" });
                }

                // 限制最大长度（防止超长请求）
                const int maxLength = 100000; // 10万字符
                if (fullText.Length > maxLength)
                {
                    return BadRequest(new { error = $"文档过长（{fullText.Length} 字符），最大支持 {maxLength} 字符" });
                }

                // 执行校对（支持客户端取消）
                var stopwatch = Stopwatch.StartNew();
                var results = await proofreadService.ProofreadDocumentAsync(fullText, HttpContext.RequestAborted);
                stopwatch.Stop();

                // 转换为响应格式
                var response = ConvertToResponse(results, request.Paragraphs);
                var failedCount = results.Count(r => r != null && !r.IsCompleted);

                _logger.LogInformation("Proofread completed in {ElapsedMs}ms, found {IssueCount} issues, {FailedCount} paragraphs failed",
                    stopwatch.ElapsedMilliseconds, response.Count, failedCount);

                var cacheStats = ProofreadService.GetCacheStats();
                return Ok(new ProofreadResponse
                {
                    Success = true,
                    Issues = response,
                    TotalParagraphs = results.Count,
                    ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
                    FailedParagraphs = failedCount,
                    CacheStats = new CacheStatsInfo { Count = cacheStats.count, TotalBytes = cacheStats.totalBytes }
                });
            }
            catch (LLMServiceException ex)
            {
                _logger.LogError(ex, "LLM service error during proofread");
                return StatusCode(502, new { error = ex.GetFriendlyErrorMessage() });
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("Proofread was cancelled");
                return StatusCode(499, new { error = "请求被取消" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Proofread failed");
                return StatusCode(500, new { error = $"校对失败: {ex.Message}" });
            }
        }

        [HttpPost("cancel")]
        [EnableRateLimiting("ProofreadPolicy")]
        public IActionResult Cancel()
        {
            try
            {
                ProofreadService.CancelAll();
                _logger.LogInformation("Global proofread cancellation requested");
                return Ok(new { success = true, message = "已发送取消信号" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Cancel request failed");
                return StatusCode(500, new { error = $"取消失败: {ex.Message}" });
            }
        }

        [HttpGet("health")]
        public IActionResult Health()
        {
            return Ok(new { 
                status = "ok", 
                timestamp = DateTime.UtcNow,
                version = "1.0.0"
            });
        }

        [HttpGet("config")]
        public IActionResult GetConfig()
        {
            var config = ConfigManager.CurrentConfig;
            return Ok(new
            {
                provider = config.Provider.ToString(),
                model = config.Model,
                proofreadMode = config.ProofreadMode,
                autoConnect = config.AutoConnect,
                // 不返回 API Key（安全）
                hasApiKey = !string.IsNullOrEmpty(config.ApiKey)
            });
        }

        [HttpPost("config")]
        [EnableRateLimiting("ProofreadPolicy")]
        public IActionResult SaveConfig([FromBody] SaveConfigRequest request)
        {
            try
            {
                if (request == null)
                {
                    return BadRequest(new { error = "请求体不能为空" });
                }

                if (!Enum.TryParse<AIProvider>(request.Provider, out var provider))
                {
                    return BadRequest(new { error = "无效的提供商" });
                }

                ConfigManager.SaveCurrentConfig(
                    provider, 
                    request.ApiKey, 
                    request.ApiUrl, 
                    request.Model, 
                    request.AutoConnect
                );

                _logger.LogInformation("Configuration saved for provider {Provider}", request.Provider);

                return Ok(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Save config failed");
                return StatusCode(500, new { error = $"保存配置失败: {ex.Message}" });
            }
        }

        [HttpPost("clear-cache")]
        [EnableRateLimiting("ProofreadPolicy")]
        public IActionResult ClearCache()
        {
            try
            {
                ProofreadService.ClearCache();
                _logger.LogInformation("Cache cleared");
                return Ok(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Clear cache failed");
                return StatusCode(500, new { error = $"清空缓存失败: {ex.Message}" });
            }
        }

        [HttpGet("stats")]
        public IActionResult GetStats()
        {
            var stats = ProofreadService.GetCacheStats();
            return Ok(new
            {
                cacheCount = stats.count,
                cacheBytes = stats.totalBytes,
                cacheBytesFormatted = $"{stats.totalBytes / 1024.0:F1} KB"
            });
        }

        /// <summary>
        /// 转换校对结果为响应格式，计算全局偏移量
        /// </summary>
        private List<ProofreadIssue> ConvertToResponse(List<ParagraphResult> results, List<ParagraphInfo> paragraphs)
        {
            var issues = new List<ProofreadIssue>();
            int issueIndex = 1;

            foreach (var result in results)
            {
                if (result == null) continue;
                if (result.Index < 0 || result.Index >= paragraphs.Count) continue;
                if (!result.IsCompleted) continue; // 跳过失败的段落

                var para = paragraphs[result.Index];
                var items = result.Items;

                foreach (var item in items)
                {
                    if (string.IsNullOrWhiteSpace(item.Original)) continue;

                    // 使用改进的偏移量计算方法
                    var localOffset = FindTextOffset(para.Text, item.Original);
                    if (localOffset < 0) localOffset = 0;

                    // 计算全局偏移量
                    var startOffset = para.StartOffset + localOffset;
                    var endOffset = startOffset + item.Original.Length;

                    // 验证偏移量范围
                    if (endOffset > para.EndOffset)
                    {
                        _logger.LogWarning(
                            "Offset out of range for issue {Index}: start={Start}, end={End}, paraEnd={ParaEnd}",
                            issueIndex, startOffset, endOffset, para.EndOffset);
                        endOffset = Math.Min(endOffset, para.EndOffset);
                    }

                    issues.Add(new ProofreadIssue
                    {
                        Index = issueIndex++,
                        ParagraphIndex = result.Index,
                        StartOffset = startOffset,
                        EndOffset = endOffset,
                        Original = item.Original,
                        Suggestion = item.Modified,
                        Reason = item.Reason,
                        Severity = item.Severity ?? "medium",
                        Type = item.Type ?? "其他"
                    });
                }
            }

            return issues;
        }

        /// <summary>
        /// 在段落中查找文本位置，使用模糊匹配提高准确性
        /// </summary>
        private int FindTextOffset(string paragraph, string text)
        {
            if (string.IsNullOrEmpty(paragraph) || string.IsNullOrEmpty(text))
                return -1;

            // 精确匹配
            var exactIndex = paragraph.IndexOf(text, StringComparison.Ordinal);
            if (exactIndex >= 0)
                return exactIndex;

            // 如果不精确匹配，尝试忽略空白字符的匹配
            var normalizedPara = NormalizeText(paragraph);
            var normalizedText = NormalizeText(text);
            
            if (normalizedPara.Contains(normalizedText))
            {
                // 找到归一化后的位置，再映射回原文
                var normIndex = normalizedPara.IndexOf(normalizedText, StringComparison.Ordinal);
                return MapNormalizedIndexToOriginal(paragraph, normIndex);
            }

            return -1;
        }

        /// <summary>
        /// 归一化文本（去除多余空白）
        /// </summary>
        private string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            
            // 替换多种空白为单个空格
            var normalized = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
            return normalized.Trim();
        }

        /// <summary>
        /// 将归一化后的索引映射回原始文本索引
        /// </summary>
        private int MapNormalizedIndexToOriginal(string original, int normalizedIndex)
        {
            if (normalizedIndex <= 0) return 0;
            
            int originalIndex = 0;
            int normCount = 0;
            bool inWhitespace = false;
            
            foreach (char c in original)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!inWhitespace)
                    {
                        if (normCount >= normalizedIndex)
                            return originalIndex;
                        normCount++;
                        inWhitespace = true;
                    }
                }
                else
                {
                    inWhitespace = false;
                    if (normCount >= normalizedIndex)
                        return originalIndex;
                    normCount++;
                }
                originalIndex++;
            }
            
            return original.Length;
        }
    }

    // API 请求/响应模型
    public class ProofreadRequest
    {
        public List<ParagraphInfo> Paragraphs { get; set; } = new();
        public string Provider { get; set; } = "DeepSeek";
        public string ApiKey { get; set; } = "";
        public string ApiUrl { get; set; } = "";
        public string Model { get; set; } = "";
        public string Prompt { get; set; } = "";
        public ProofreadMode Mode { get; set; } = ProofreadMode.Precise;
    }

    public class ParagraphInfo
    {
        [Newtonsoft.Json.JsonProperty("index")]
        public int Index { get; set; }
        
        [Newtonsoft.Json.JsonProperty("start")]
        public int StartOffset { get; set; }
        
        [Newtonsoft.Json.JsonProperty("end")]
        public int EndOffset { get; set; }
        
        [Newtonsoft.Json.JsonProperty("text")]
        public string Text { get; set; } = "";
    }

    public class ProofreadIssue
    {
        [Newtonsoft.Json.JsonProperty("index")]
        public int Index { get; set; }
        
        [Newtonsoft.Json.JsonProperty("paragraphIndex")]
        public int ParagraphIndex { get; set; }
        
        [Newtonsoft.Json.JsonProperty("startOffset")]
        public int StartOffset { get; set; }
        
        [Newtonsoft.Json.JsonProperty("endOffset")]
        public int EndOffset { get; set; }
        
        [Newtonsoft.Json.JsonProperty("original")]
        public string Original { get; set; } = "";
        
        [Newtonsoft.Json.JsonProperty("suggestion")]
        public string Suggestion { get; set; } = "";
        
        [Newtonsoft.Json.JsonProperty("reason")]
        public string Reason { get; set; } = "";
        
        [Newtonsoft.Json.JsonProperty("severity")]
        public string Severity { get; set; } = "";
        
        [Newtonsoft.Json.JsonProperty("type")]
        public string Type { get; set; } = "";
    }

    public class CacheStatsInfo
    {
        [Newtonsoft.Json.JsonProperty("count")]
        public int Count { get; set; }
        
        [Newtonsoft.Json.JsonProperty("totalBytes")]
        public long TotalBytes { get; set; }
    }

    public class ProofreadResponse
    {
        [Newtonsoft.Json.JsonProperty("success")]
        public bool Success { get; set; }
        
        [Newtonsoft.Json.JsonProperty("issues")]
        public List<ProofreadIssue> Issues { get; set; } = new();
        
        [Newtonsoft.Json.JsonProperty("totalParagraphs")]
        public int TotalParagraphs { get; set; }
        
        [Newtonsoft.Json.JsonProperty("elapsedSeconds")]
        public double ElapsedSeconds { get; set; }

        [Newtonsoft.Json.JsonProperty("failedParagraphs")]
        public int FailedParagraphs { get; set; }
        
        [Newtonsoft.Json.JsonProperty("cacheStats")]
        public CacheStatsInfo CacheStats { get; set; } = new CacheStatsInfo();
    }

    public enum ProofreadMode
    {
        Precise,
        FullText
    }

    public class SaveConfigRequest
    {
        public string Provider { get; set; } = "";
        public string ApiKey { get; set; } = "";
        public string ApiUrl { get; set; } = "";
        public string Model { get; set; } = "";
        public bool AutoConnect { get; set; }
    }
}
