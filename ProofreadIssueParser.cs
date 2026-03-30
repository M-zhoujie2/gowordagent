using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 校对结果解析器 - 解析 AI 返回的校对结果
    /// </summary>
    public static class ProofreadIssueParser
    {
        // 预编译的正则表达式（提高性能）
        private static readonly Regex ProofreadItemRegex = new Regex(
            @"【第(?<index>\d+)处】类型：(?<type>[^\r\n|]+)(?:[｜|]严重度：(?<severity>[^\r\n]+))?\r?\n原文：(?<original>.*?)\r?\n修改：(?<modified>.*?)\r?\n理由：(?<reason>.*?)(?=\r?\n【第|$)", 
            RegexOptions.Singleline | RegexOptions.Compiled);

        /// <summary>
        /// 解析 AI 返回的校对结果文本
        /// </summary>
        public static List<ProofreadIssueItem> ParseProofreadItems(string aiResponse)
        {
            var items = new List<ProofreadIssueItem>();
            
            if (string.IsNullOrWhiteSpace(aiResponse))
                return items;
            
            int idx = 1;
            foreach (Match match in ProofreadItemRegex.Matches(aiResponse))
            {
                string original = match.Groups["original"].Value.Trim();
                if (!string.IsNullOrWhiteSpace(original))
                {
                    items.Add(new ProofreadIssueItem
                    {
                        Index = idx++,
                        Type = match.Groups["type"].Value.Trim(),
                        Severity = match.Groups["severity"].Value.Trim().ToLower(),
                        Original = original,
                        Modified = match.Groups["modified"].Value.Trim(),
                        Reason = match.Groups["reason"].Value.Trim()
                    });
                }
            }
            return items;
        }

        /// <summary>
        /// 统计问题数量
        /// </summary>
        public static int CountIssues(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            var matches = Regex.Matches(text, @"【第\d+处】");
            return matches.Count;
        }

        /// <summary>
        /// 按类型统计问题
        /// </summary>
        public static Dictionary<string, int> CategorizeIssues(string text)
        {
            var categories = new Dictionary<string, int>();
            if (string.IsNullOrWhiteSpace(text)) return categories;

            var pattern = @"【第\d+处】类型：([^\r\n:]+)";
            foreach (Match match in Regex.Matches(text, pattern))
            {
                string cat = match.Groups[1].Value.Trim();
                
                // 归一化分类
                if (cat.Contains("错别字") || cat.Contains("拼写"))
                    cat = "错别字";
                else if (cat.Contains("语病") || cat.Contains("语法") || cat.Contains("搭配") || cat.Contains("杂糅"))
                    cat = "语病";
                else if (cat.Contains("标点"))
                    cat = "标点错误";
                else if (cat.Contains("序号"))
                    cat = "序号问题";
                else if (cat.Contains("用词"))
                    cat = "用词不当";
                else if (cat.Contains("术语") || cat.Contains("不一致"))
                    cat = "术语不一致";
                else if (cat.Contains("格式"))
                    cat = "格式问题";
                else if (cat.Contains("逻辑") || cat.Contains("矛盾"))
                    cat = "逻辑/矛盾";
                else
                    cat = "其他";

                if (categories.ContainsKey(cat))
                    categories[cat]++;
                else
                    categories[cat] = 1;
            }

            return categories;
        }
    }
}
