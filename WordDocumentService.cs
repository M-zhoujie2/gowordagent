using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// Word 文档操作服务，封装所有 Word COM 交互
    /// </summary>
    public class WordDocumentService : IDisposable
    {
        private readonly Word.Application _application;
        private readonly Word.Document _document;
        private bool _disposed;

        public WordDocumentService(Word.Application application, Word.Document document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// 获取当前活动文档的内容或选中的文本
        /// </summary>
        public static string GetDocumentText(Word.Application app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            Word.Document doc = null;
            Word.Window activeWindow = null;
            Word.Selection selection = null;
            Word.Range selectedRange = null;
            Word.Range contentRange = null;

            try
            {
                doc = app.ActiveDocument;
                if (doc == null) throw new InvalidOperationException("当前未打开任何文档。");

                string selectedText = string.Empty;
                
                // 安全获取选中文本
                try
                {
                    activeWindow = doc.ActiveWindow;
                    if (activeWindow != null)
                    {
                        selection = activeWindow.Selection;
                        if (selection != null)
                        {
                            selectedRange = selection.Range;
                            if (selectedRange != null)
                            {
                                selectedText = selectedRange.Text ?? string.Empty;
                            }
                        }
                    }
                }
                catch (COMException)
                {
                    selectedText = string.Empty;
                }

                string docText;
                if (!string.IsNullOrWhiteSpace(selectedText))
                {
                    docText = selectedText;
                }
                else
                {
                    contentRange = doc.Content;
                    docText = contentRange?.Text ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(docText))
                    throw new InvalidOperationException("文档正文为空。");

                return docText;
            }
            finally
            {
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
                if (selection != null) Marshal.ReleaseComObject(selection);
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                if (contentRange != null) Marshal.ReleaseComObject(contentRange);
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
        }

        /// <summary>
        /// 检查文档是否有效（未被释放）
        /// </summary>
        public bool IsDocumentValid()
        {
            Word.Range content = null;
            try
            {
                content = _document.Content;
                var _ = content.Text;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (content != null) Marshal.ReleaseComObject(content);
            }
        }

        /// <summary>
        /// 应用单个修订到文档
        /// </summary>
        public bool ApplyRevision(string original, string modified, string commentText, out int start, out int end)
        {
            start = -1;
            end = -1;

            if (!IsDocumentValid()) return false;

            Word.Range searchRange = null;
            Word.Find find = null;
            Word.Comment comment = null;

            try
            {
                searchRange = _document.Content;
                find = searchRange.Find;

                bool found = find.Execute(FindText: original, MatchCase: false, MatchWholeWord: false,
                                          MatchWildcards: false, Forward: true, Wrap: Word.WdFindWrap.wdFindStop);

                if (found)
                {
                    start = searchRange.Start;
                    end = searchRange.End;

                    bool oldTrackRevisions = _document.TrackRevisions;
                    try
                    {
                        _document.TrackRevisions = true;
                        searchRange.Text = modified;
                        comment = _document.Comments.Add(searchRange, commentText);
                        end = searchRange.End;
                        return true;
                    }
                    finally
                    {
                        _document.TrackRevisions = oldTrackRevisions;
                    }
                }
            }
            finally
            {
                if (comment != null) Marshal.ReleaseComObject(comment);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }

            return false;
        }

        /// <summary>
        /// 应用多个修订到文档
        /// </summary>
        public List<ProofreadIssueItem> ApplyRevisions(List<ProofreadIssueItem> items)
        {
            var processedItems = new List<ProofreadIssueItem>();
            if (!IsDocumentValid()) return processedItems;

            foreach (var item in items)
            {
                if (item == null) continue;
                
                try
                {
                    var commentBuilder = new StringBuilder();
                    commentBuilder.AppendLine($"【第{item.Index}处】类型：{item.Type}{(string.IsNullOrEmpty(item.Severity) ? "" : $"｜严重度：{item.Severity}")}");
                    commentBuilder.AppendLine($"原文：{item.Original}");
                    commentBuilder.AppendLine($"修改：{item.Modified}");
                    commentBuilder.AppendLine($"理由：{item.Reason}");

                    if (ApplyRevision(item.Original, item.Modified, commentBuilder.ToString(), out int start, out int end))
                    {
                        item.DocumentStart = start;
                        item.DocumentEnd = end;
                        processedItems.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyRevisions] 处理项目时出错: {ex.Message}");
                }
            }

            return processedItems;
        }

        /// <summary>
        /// 导航到指定位置
        /// </summary>
        public bool NavigateToRange(int start, int end)
        {
            if (!IsDocumentValid()) return false;
            if (start < 0 || end <= start) return false;

            Word.Range range = null;
            Word.Window activeWindow = null;
            try
            {
                range = _document.Range(start, end);
                range.Select();
                activeWindow = _application.ActiveWindow;
                activeWindow.ScrollIntoView(range);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToRange] 使用位置定位失败: {ex.Message}");
                return false;
            }
            finally
            {
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                if (range != null) Marshal.ReleaseComObject(range);
            }
        }

        /// <summary>
        /// 通过原文搜索定位
        /// </summary>
        public bool NavigateBySearch(string originalText)
        {
            if (!IsDocumentValid()) return false;
            if (string.IsNullOrWhiteSpace(originalText)) return false;

            Word.Range searchRange = null;
            Word.Find find = null;
            Word.Window activeWindow = null;
            try
            {
                searchRange = _document.Content;
                find = searchRange.Find;
                
                if (find.Execute(FindText: originalText, MatchCase: false, MatchWholeWord: false,
                                  MatchWildcards: false, Forward: true, Wrap: Word.WdFindWrap.wdFindStop))
                {
                    searchRange.Select();
                    activeWindow = _application.ActiveWindow;
                    activeWindow.ScrollIntoView(searchRange);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateBySearch] 搜索定位失败: {ex.Message}");
            }
            finally
            {
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }

            return false;
        }

        /// <summary>
        /// 导航到校对问题项
        /// </summary>
        public void NavigateToIssue(ProofreadIssueItem item)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));
            if (_application == null) throw new InvalidOperationException("Application 为空");
            if (_document == null) throw new InvalidOperationException("Document 为空");
            if (!IsDocumentValid()) throw new InvalidOperationException("文档已被释放");

            // 激活 Word 窗口
            try 
            { 
                _application.Activate(); 
            } 
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToIssue] 激活窗口失败: {ex.Message}");
            }

            // 优先使用位置定位
            if (NavigateToRange(item.DocumentStart, item.DocumentEnd))
                return;

            // 备用方案：通过原文搜索
            if (!NavigateBySearch(item.Original))
                throw new InvalidOperationException("无法在文档中找到该位置，可能文本已被修改。");
        }

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // 注意：_application 和 _document 是外部传入的引用，
                // 不应该在这里释放，否则会影响 Word 的正常运行
                _disposed = true;
            }
        }

        #endregion
    }

    /// <summary>
    /// Word 文档服务工厂
    /// </summary>
    public static class WordDocumentServiceFactory
    {
        /// <summary>
        /// 尝试创建文档服务
        /// </summary>
        public static bool TryCreate(out WordDocumentService service, out string errorMessage)
        {
            service = null;
            errorMessage = null;

            Word.Document doc = null;

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    errorMessage = "无法访问 Word 应用";
                    return false;
                }

                doc = app.ActiveDocument;
                if (doc == null)
                {
                    errorMessage = "当前未打开任何文档";
                    return false;
                }

                // 检查文档是否有效
                try
                {
                    var _ = doc.Content.Text;
                }
                catch (COMException)
                {
                    Marshal.ReleaseComObject(doc);
                    errorMessage = "文档已被释放，请重新打开";
                    return false;
                }

                service = new WordDocumentService(app, doc);
                return true;
            }
            catch (Exception ex)
            {
                if (doc != null) Marshal.ReleaseComObject(doc);
                errorMessage = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 创建文档服务（失败时抛出异常）
        /// </summary>
        public static WordDocumentService Create()
        {
            if (!TryCreate(out var service, out var errorMessage))
                throw new InvalidOperationException(errorMessage);
            return service;
        }
    }
}
