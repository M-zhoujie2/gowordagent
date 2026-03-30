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
    /// 支持多文档场景：绑定特定文档，防止切换文档导致引用失效
    /// </summary>
    public class WordProofreadController : IDisposable
    {
        private readonly Dispatcher _dispatcher;
        private WordDocumentService _documentService;
        private Word.Document _boundDocument;
        private readonly object _lock = new object();
        private bool _disposed;

        public WordProofreadController(Dispatcher dispatcher = null)
        {
            _dispatcher = dispatcher ?? Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
        }

        /// <summary>
        /// 获取当前绑定的文档服务（多文档安全）
        /// </summary>
        private bool TryGetDocumentService(out WordDocumentService service, out string errorMessage)
        {
            lock (_lock)
            {
                // 检查当前活动文档是否与绑定文档一致
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    service = null;
                    errorMessage = "无法访问 Word 应用";
                    return false;
                }

                var activeDoc = app.ActiveDocument;
                if (activeDoc == null)
                {
                    service = null;
                    errorMessage = "当前未打开任何文档";
                    return false;
                }

                // 如果尚未绑定或绑定的是不同文档，重新创建服务
                if (_documentService == null || !IsSameDocument(_boundDocument, activeDoc))
                {
                    // 释放旧服务
                    _documentService?.Dispose();
                    _documentService = null;
                    _boundDocument = null;

                    // 创建新服务
                    if (!WordDocumentServiceFactory.TryCreateForDocument(app, activeDoc, out _documentService, out errorMessage))
                    {
                        service = null;
                        return false;
                    }
                    _boundDocument = activeDoc;
                    System.Diagnostics.Debug.WriteLine($"[WordProofreadController] 绑定到新文档: {activeDoc.Name}");
                }

                // 验证文档是否仍有效
                if (!_documentService.IsDocumentValid())
                {
                    _documentService.Dispose();
                    _documentService = null;
                    _boundDocument = null;
                    errorMessage = "文档已被关闭或释放";
                    service = null;
                    return false;
                }

                service = _documentService;
                errorMessage = null;
                return true;
            }
        }

        /// <summary>
        /// 检查两个文档引用是否指向同一文档
        /// </summary>
        private bool IsSameDocument(Word.Document doc1, Word.Document doc2)
        {
            if (doc1 == null || doc2 == null) return false;
            try
            {
                return doc1.FullName == doc2.FullName;
            }
            catch
            {
                // 如果访问属性失败，假设不是同一文档
                return false;
            }
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
        /// 支持多文档：绑定到特定文档，切换文档时会重新绑定
        /// </summary>
        public List<ProofreadIssueItem> ApplyProofreadToDocument(List<ProofreadIssueItem> items, Action<string, string, bool, bool> addMessageCallback = null)
        {
            try
            {
                return _dispatcher.Invoke(() =>
                {
                    if (!TryGetDocumentService(out var service, out var errorMessage))
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordProofreadController] {errorMessage}");
                        addMessageCallback?.Invoke("错误", errorMessage, false, true);
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
        /// 支持多文档：绑定到特定文档，切换文档时会重新绑定
        /// </summary>
        public void NavigateToIssue(ProofreadIssueItem item)
        {
            try
            {
                _dispatcher.Invoke(() =>
                {
                    try
                    {
                        if (!TryGetDocumentService(out var service, out var errorMessage))
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

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                lock (_lock)
                {
                    _documentService?.Dispose();
                    _documentService = null;
                    _boundDocument = null;
                }
                _disposed = true;
            }
        }

        #endregion
    }
}
