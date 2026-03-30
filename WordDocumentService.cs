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
    public class WordDocumentService
    {
        private readonly Word.Application _application;
        private readonly Word.Document _document;

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

            var doc = app.ActiveDocument;
            if (doc == null) throw new InvalidOperationException("当前未打开任何文档。");

            string selectedText = doc.ActiveWindow?.Selection?.Range?.Text ?? string.Empty;
            string docText = !string.IsNullOrWhiteSpace(selectedText) 
                ? selectedText 
                : (doc.Content?.Text ?? string.Empty);

            if (string.IsNullOrWhiteSpace(docText))
                throw new InvalidOperationException("文档正文为空。");

            return docText;
        }

        /// <summary>
        /// 检查文档是否有效（未被释放）
        /// </summary>
        public bool IsDocumentValid()
        {
            try
            {
                var _ = _document.Content.Text;
                return true;
            }
            catch (COMException)
            {
                return false;
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
            try
            {
                searchRange = _document.Content;
                if (searchRange.Find.Execute(FindText: original, MatchCase: false, MatchWholeWord: false,
                                              MatchWildcards: false, Forward: true, Wrap: Word.WdFindWrap.wdFindStop))
                {
                    start = searchRange.Start;
                    end = searchRange.End;

                    bool oldTrackRevisions = _document.TrackRevisions;
                    try
                    {
                        _document.TrackRevisions = true;
                        searchRange.Text = modified;
                        _document.Comments.Add(searchRange, commentText);
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
                if (searchRange != null)
                    Marshal.ReleaseComObject(searchRange);
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
            try
            {
                range = _document.Range(start, end);
                range.Select();
                _application.ActiveWindow.ScrollIntoView(range);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToRange] 使用位置定位失败: {ex.Message}");
                return false;
            }
            finally
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);
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
            try
            {
                searchRange = _document.Content;
                if (searchRange.Find.Execute(FindText: originalText, MatchCase: false, MatchWholeWord: false,
                                              MatchWildcards: false, Forward: true, Wrap: Word.WdFindWrap.wdFindStop))
                {
                    searchRange.Select();
                    _application.ActiveWindow.ScrollIntoView(searchRange);
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateBySearch] 搜索定位失败: {ex.Message}");
            }
            finally
            {
                if (searchRange != null)
                    Marshal.ReleaseComObject(searchRange);
            }

            return false;
        }

        /// <summary>
        /// 导航到校对问题项
        /// </summary>
        public void NavigateToIssue(ProofreadIssueItem item)
        {
            if (_application == null) throw new InvalidOperationException("Application 为空");
            if (_document == null) throw new InvalidOperationException("Document 为空");
            if (!IsDocumentValid()) throw new InvalidOperationException("文档已被释放");

            // 激活 Word 窗口
            try { _application.Activate(); } catch { }

            // 优先使用位置定位
            if (NavigateToRange(item.DocumentStart, item.DocumentEnd))
                return;

            // 备用方案：通过原文搜索
            if (!NavigateBySearch(item.Original))
                throw new InvalidOperationException("无法在文档中找到该位置，可能文本已被修改。");
        }
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

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    errorMessage = "无法访问 Word 应用";
                    return false;
                }

                var doc = app.ActiveDocument;
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
                    errorMessage = "文档已被释放，请重新打开";
                    return false;
                }

                service = new WordDocumentService(app, doc);
                return true;
            }
            catch (Exception ex)
            {
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
