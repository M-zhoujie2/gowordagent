using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 文档分段配置
    /// </summary>
    public class SegmenterConfig
    {
        private int _targetChunkSize = 1500;
        private int _overlapSize = 100;
        
        /// <summary>
        /// 目标每段大小（字符数）
        /// </summary>
        public int TargetChunkSize 
        { 
            get => _targetChunkSize;
            set => _targetChunkSize = Math.Max(100, value); // 最小100字符
        }

        /// <summary>
        /// 重叠大小（字符数）
        /// </summary>
        public int OverlapSize 
        { 
            get => _overlapSize;
            set => _overlapSize = Math.Max(0, Math.Min(value, TargetChunkSize / 2)); // 不超过目标大小的一半
        }

        /// <summary>
        /// 句子分隔正则表达式
        /// </summary>
        public string SentencePattern { get; set; } = @"(?<=[。！？.!?；;])";
    }

    /// <summary>
    /// 文档分段器 - 将长文档分割成适合 LLM 处理的段落
    /// </summary>
    public class DocumentSegmenter
    {
        private readonly SegmenterConfig _config;

        public DocumentSegmenter(SegmenterConfig config = null)
        {
            _config = config ?? new SegmenterConfig();
        }

        /// <summary>
        /// 将文档分段
        /// </summary>
        public List<string> SplitIntoParagraphs(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<string>();

            // 如果文本较短，直接作为一个段落
            if (text.Length <= _config.TargetChunkSize)
            {
                Debug.WriteLine($"[DocumentSegmenter] 文本较短({text.Length}字)，不分段");
                return new List<string> { text.Trim() };
            }

            var chunks = new List<string>();
            var sentences = SplitIntoSentences(text);
            var currentChunk = new StringBuilder();

            for (int i = 0; i < sentences.Count; i++)
            {
                var sentence = sentences[i];

                // 添加句子到当前块
                currentChunk.Append(sentence);

                // 检查是否需要切分（达到目标大小）
                if (currentChunk.Length >= _config.TargetChunkSize)
                {
                    // 达到目标块大小，准备切分
                    var chunkText = currentChunk.ToString().Trim();
                    chunks.Add(chunkText);

                    // 计算重叠内容（从块末尾提取）
                    var overlapText = GetOverlapText(chunkText, _config.OverlapSize);

                    // 开始新块，带上重叠内容
                    currentChunk.Clear();
                    currentChunk.Append(overlapText);
                }
            }

            // 处理剩余内容
            if (currentChunk.Length > 0)
            {
                chunks.Add(currentChunk.ToString().Trim());
            }

            Debug.WriteLine($"[DocumentSegmenter] 分段完成，共 {chunks.Count} 段");

            return chunks;
        }

        /// <summary>
        /// 将文本分割成句子列表
        /// </summary>
        public List<string> SplitIntoSentences(string text)
        {
            if (string.IsNullOrEmpty(text))
                return new List<string>();
            
            if (string.IsNullOrEmpty(_config.SentencePattern))
                return new List<string> { text };

            try
            {
                var sentences = Regex.Split(text, _config.SentencePattern)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                return sentences;
            }
            catch (ArgumentException ex)
            {
                // 正则表达式无效，返回整个文本作为一个句子
                Debug.WriteLine($"[DocumentSegmenter] 正则表达式无效: {ex.Message}");
                return new List<string> { text };
            }
        }

        /// <summary>
        /// 获取重叠文本（从块末尾提取，用于下一块开头）
        /// </summary>
        private string GetOverlapText(string text, int overlapSize)
        {
            if (text.Length <= overlapSize)
                return text;

            return text.Substring(text.Length - overlapSize);
        }
    }
}
