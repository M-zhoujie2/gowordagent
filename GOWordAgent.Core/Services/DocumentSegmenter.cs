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

        public int TargetChunkSize
        {
            get => _targetChunkSize;
            set
            {
                _targetChunkSize = Math.Max(100, value);
                _overlapSize = Math.Min(_overlapSize, _targetChunkSize / 2);
            }
        }

        public int OverlapSize
        {
            get => _overlapSize;
            set => _overlapSize = Math.Max(0, Math.Min(value, _targetChunkSize / 2));
        }

        public string SentencePattern { get; set; } = @"(?<=[。！？.!?；;])";
    }

    /// <summary>
    /// 文档分段器
    /// </summary>
    public class DocumentSegmenter
    {
        private readonly SegmenterConfig _config;

        public DocumentSegmenter(SegmenterConfig? config = null)
        {
            _config = config ?? new SegmenterConfig();
        }

        public List<string> SplitIntoParagraphs(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<string>();

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
                currentChunk.Append(sentence);

                if (currentChunk.Length >= _config.TargetChunkSize)
                {
                    var chunkText = currentChunk.ToString().Trim();
                    chunks.Add(chunkText);

                    var overlapText = GetOverlapText(chunkText, _config.OverlapSize);
                    currentChunk.Clear();
                    currentChunk.Append(overlapText);
                }
            }

            if (currentChunk.Length > 0)
            {
                chunks.Add(currentChunk.ToString().Trim());
            }

            Debug.WriteLine($"[DocumentSegmenter] 分段完成，共 {chunks.Count} 段");
            return chunks;
        }

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
                Debug.WriteLine($"[DocumentSegmenter] 正则表达式无效: {ex.Message}");
                return new List<string> { text };
            }
        }

        private string GetOverlapText(string text, int overlapSize)
        {
            if (text.Length <= overlapSize)
                return text;

            return text.Substring(text.Length - overlapSize);
        }
    }
}
