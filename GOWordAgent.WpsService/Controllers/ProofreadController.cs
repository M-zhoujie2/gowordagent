using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using GOWordAgentAddIn;
using GOWordAgentAddIn.Models;

namespace GOWordAgent.WpsService.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ProofreadController : ControllerBase
    {
        private readonly ILLMService _llmService;
        private readonly ILogger<ProofreadController> _logger;

        public ProofreadController(ILLMService llmService, ILogger<ProofreadController> logger)
        {
            _llmService = llmService;
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> Proofread([FromBody] ProofreadRequest request)
        {
            try
            {
                _logger.LogInformation("Starting proofread for document with {ParagraphCount} paragraphs", 
                    request.Paragraphs?.Count ?? 0);

                // 获取提示词
                var prompt = string.IsNullOrEmpty(request.Prompt) 
                    ? ConfigManager.GetProofreadPromptForMode(request.Mode.ToString())
                    : request.Prompt;

                // 创建校对服务
                var proofreadService = new ProofreadService(
                    _llmService,
                    prompt,
                    concurrency: 5,
                    proofreadMode: request.Mode.ToString()
                );

                // 合并段落文本
                var fullText = string.Join("\n", request.Paragraphs?.Select(p => p.Text) ?? Array.Empty<string>());

                if (string.IsNullOrWhiteSpace(fullText))
                {
                    return BadRequest(new { error = "文档内容为空" });
                }

                // 执行校对
                var stopwatch = Stopwatch.StartNew();
                var results = await proofreadService.ProofreadDocumentAsync(fullText);
                stopwatch.Stop();

                // 转换为响应格式
                var response = ConvertToResponse(results, request.Paragraphs ?? new List<ParagraphInfo>());

                _logger.LogInformation("Proofread completed in {ElapsedMs}ms, found {IssueCount} issues",
                    stopwatch.ElapsedMilliseconds, response.Count);

                return Ok(new ProofreadResponse
                {
                    Success = true,
                    Issues = response,
                    TotalParagraphs = results.Count,
                    ElapsedSeconds = stopwatch.Elapsed.TotalSeconds,
                    CacheStats = ProofreadService.GetCacheStats()
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Proofread failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpGet("health")]
        public IActionResult Health()
        {
            return Ok(new { status = "ok", timestamp = DateTime.UtcNow });
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
                autoConnect = config.AutoConnect
            });
        }

        [HttpPost("config")]
        public IActionResult SaveConfig([FromBody] SaveConfigRequest request)
        {
            try
            {
                Enum.TryParse<AIProvider>(request.Provider, out var provider);
                ConfigManager.SaveCurrentConfig(provider, request.ApiKey, request.ApiUrl, request.Model, request.AutoConnect);
                return Ok(new { success = true });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Save config failed");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        private List<ProofreadIssue> ConvertToResponse(List<ParagraphResult> results, List<ParagraphInfo> paragraphs)
        {
            var issues = new List<ProofreadIssue>();

            foreach (var result in results)
            {
                if (result.Index >= paragraphs.Count) continue;

                var para = paragraphs[result.Index];
                var items = ProofreadIssueParser.ParseProofreadItems(result.ResultText);

                foreach (var item in items)
                {
                    // 计算在段落内的偏移量
                    var localOffset = FindTextOffset(para.Text, item.Original);
                    if (localOffset < 0) localOffset = 0;

                    // 计算全局偏移量
                    var startOffset = para.StartOffset + localOffset;
                    var endOffset = startOffset + item.Original.Length;

                    issues.Add(new ProofreadIssue
                    {
                        ParagraphIndex = result.Index,
                        StartOffset = startOffset,
                        EndOffset = endOffset,
                        Original = item.Original,
                        Suggestion = item.Modified,
                        Reason = item.Reason,
                        Severity = item.Severity,
                        Type = item.Type
                    });
                }
            }

            return issues;
        }

        private int FindTextOffset(string paragraph, string text)
        {
            if (string.IsNullOrEmpty(paragraph) || string.IsNullOrEmpty(text))
                return 0;

            var index = paragraph.IndexOf(text, StringComparison.Ordinal);
            return index >= 0 ? index : 0;
        }
    }

    // API 请求/响应模型
    public class ProofreadRequest
    {
        public string Text { get; set; } = "";
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
        public int Index { get; set; }
        public int StartOffset { get; set; }
        public int EndOffset { get; set; }
        public string Text { get; set; } = "";
    }

    public class ProofreadIssue
    {
        public int ParagraphIndex { get; set; }
        public int StartOffset { get; set; }
        public int EndOffset { get; set; }
        public string Original { get; set; } = "";
        public string Suggestion { get; set; } = "";
        public string Reason { get; set; } = "";
        public string Severity { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ProofreadResponse
    {
        public bool Success { get; set; }
        public List<ProofreadIssue> Issues { get; set; } = new();
        public int TotalParagraphs { get; set; }
        public double ElapsedSeconds { get; set; }
        public (int count, long totalBytes) CacheStats { get; set; }
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
