using System;
using System.Collections.Generic;

namespace GOWordAgentAddIn.Models
{
    /// <summary>
    /// 段落校验结果
    /// </summary>
    public class ParagraphResult
    {
        /// <summary>
        /// 段落索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 原始文本
        /// </summary>
        public string OriginalText { get; set; }

        /// <summary>
        /// 完整结果文本
        /// </summary>
        public string ResultText { get; set; }

        /// <summary>
        /// 是否已完成
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 是否来自缓存
        /// </summary>
        public bool IsCached { get; set; }

        /// <summary>
        /// 处理时间
        /// </summary>
        public DateTime ProcessTime { get; set; }

        /// <summary>
        /// 耗时（毫秒）
        /// </summary>
        public long ElapsedMs { get; set; }

        /// <summary>
        /// 解析出的校对项
        /// </summary>
        public List<ProofreadIssueItem> Items { get; set; } = new List<ProofreadIssueItem>();
    }

    /// <summary>
    /// 校验进度回调参数
    /// </summary>
    public class ProofreadProgressArgs : EventArgs
    {
        /// <summary>
        /// 总段落数
        /// </summary>
        public int TotalParagraphs { get; set; }

        /// <summary>
        /// 已完成段落数
        /// </summary>
        public int CompletedParagraphs { get; set; }

        /// <summary>
        /// 当前处理中的段落索引
        /// </summary>
        public int CurrentIndex { get; set; }

        /// <summary>
        /// 当前状态描述
        /// </summary>
        public string CurrentStatus { get; set; }

        /// <summary>
        /// 当前段落结果
        /// </summary>
        public ParagraphResult Result { get; set; }

        /// <summary>
        /// 是否全部完成
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// 预计剩余时间（秒）
        /// </summary>
        public int EstimatedRemainingSeconds { get; set; }

        /// <summary>
        /// 缓存命中数
        /// </summary>
        public int CacheHitCount { get; set; }
    }

    /// <summary>
    /// 校对问题项（带位置信息）
    /// 注意：DocumentStart/DocumentEnd 仅在应用修订时有效，
    /// 后续如果文档被修改，这些位置可能偏移。导航时应优先使用 Original 文本搜索。
    /// </summary>
    public class ProofreadIssueItem
    {
        public int Index { get; set; }
        public string Type { get; set; }
        public string Original { get; set; }
        public string Modified { get; set; }
        public string Reason { get; set; }
        public string Severity { get; set; }

        /// <summary>
        /// 在文档中的起始位置（仅作为缓存，多修订后可能偏移）
        /// </summary>
        public int DocumentStart { get; set; }

        /// <summary>
        /// 在文档中的结束位置（仅作为缓存，多修订后可能偏移）
        /// </summary>
        public int DocumentEnd { get; set; }
    }

    /// <summary>
    /// AI 服务商列表项
    /// </summary>
    public class ProviderItem
    {
        public AIProvider Provider { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }
}
