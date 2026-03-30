using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using GOWordAgentAddIn.Models;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// Word 校对文档控制器 - 处理校对结果与 Word 文档的交互
    /// </summary>
    public class WordProofreadController
    {
        private readonly Dispatcher _dispatcher;

        public WordProofreadController(Dispatcher dispatcher = null)
        {
            _dispatcher = dispatcher ?? Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
        }

        /// <summary>
        /// 获取当前文档文本
        /// </summary>
        public string GetDocumentText()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null) throw new InvalidOperationException("无法访问 Word 应用。");
                return WordDocumentService.GetDocumentText(app);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取文档失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        /// <summary>
        /// 应用校对结果到文档（批注/修订）
        /// 优化：先查找所有位置，然后倒序处理避免偏移
        /// </summary>
        public List<ProofreadIssueItem> ApplyProofreadToDocument(List<ProofreadIssueItem> items, Action<string, string, bool, bool> addMessageCallback = null)
        {
            try
            {
                return _dispatcher.Invoke(() =>
                {
                    if (!WordDocumentServiceFactory.TryCreate(out var service, out var errorMessage))
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordProofreadController] {errorMessage}");
                        return new List<ProofreadIssueItem>();
                    }

                    try
                    {
                        // 第一步：为所有问题项查找位置
                        var itemsWithPosition = FindItemPositions(service, items);
                        
                        // 第二步：按位置倒序排列（从文档末尾到开头）
                        itemsWithPosition = itemsWithPosition.OrderByDescending(x => x.start).ToList();
                        
                        // 第三步：逐个应用修订
                        var processedItems = ApplyRevisions(service, itemsWithPosition);
                        
                        // 按原始索引排序返回
                        processedItems = processedItems.OrderBy(i => i.Index).ToList();
                        
                        if (processedItems.Count > 0)
                            addMessageCallback?.Invoke("系统", $"已将 {processedItems.Count} 条诊断以批注/修订形式写入文档。", false, false);
                        
                        return processedItems;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 错误: {ex.Message}");
                        addMessageCallback?.Invoke("错误", $"写回文档时出错: {ex.Message}", false, true);
                        return new List<ProofreadIssueItem>();
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 调度错误: {ex.Message}");
                return new List<ProofreadIssueItem>();
            }
        }

        /// <summary>
        /// 在文档中定位到指定问题
        /// </summary>
        public void NavigateToIssue(ProofreadIssueItem item)
        {
            try
            {
                _dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (!WordDocumentServiceFactory.TryCreate(out var service, out var errorMessage))
                        {
                            MessageBox.Show(errorMessage, "定位失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        service.NavigateToIssue(item);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[NavigateToIssue] 错误: {ex.Message}");
                        MessageBox.Show($"定位时发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToIssue] 调度错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 构建批注文本
        /// </summary>
        public static string BuildCommentText(ProofreadIssueItem item)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"【第{item.Index}处】类型：{item.Type}{(string.IsNullOrEmpty(item.Severity) ? "" : $"｜严重度：{item.Severity}")}");
            sb.AppendLine($"原文：{item.Original}");
            sb.AppendLine($"修改：{item.Modified}");
            sb.AppendLine($"理由：{item.Reason}");
            return sb.ToString();
        }

        /// <summary>
        /// 为所有问题项查找文档位置
        /// </summary>
        private List<(ProofreadIssueItem item, int start, int end)> FindItemPositions(
            WordDocumentService service, List<ProofreadIssueItem> items)
        {
            var itemsWithPosition = new List<(ProofreadIssueItem item, int start, int end)>();
            
            foreach (var item in items)
            {
                if (item == null) continue;
                
                var (found, start, end) = service.FindTextPosition(item.Original);
                if (found)
                {
                    itemsWithPosition.Add((item, start, end));
                    System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 找到 '{item.Original.Substring(0, Math.Min(10, item.Original.Length))}...' 在位置 {start}-{end}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 未找到: '{item.Original.Substring(0, Math.Min(10, item.Original.Length))}...'");
                }
            }
            
            return itemsWithPosition;
        }

        /// <summary>
        /// 应用修订到文档
        /// </summary>
        private List<ProofreadIssueItem> ApplyRevisions(WordDocumentService service, 
            List<(ProofreadIssueItem item, int start, int end)> itemsWithPosition)
        {
            var processedItems = new List<ProofreadIssueItem>();
            
            foreach (var (item, start, end) in itemsWithPosition)
            {
                try
                {
                    if (service.ApplyRevisionAtRange(start, end, item.Original, item.Modified, 
                        BuildCommentText(item), out int newStart, out int newEnd))
                    {
                        item.DocumentStart = newStart;
                        item.DocumentEnd = newEnd;
                        processedItems.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 处理项目时出错: {ex.Message}");
                }
            }
            
            return processedItems;
        }
    }
}
