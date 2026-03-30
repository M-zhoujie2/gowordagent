using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using GOWordAgentAddIn.Models;
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
        /// 注意：此方法提取纯文本内容，表格中的文本会按文档流顺序提取，
        /// 图片、页眉页脚等内容会被忽略。如需更精细的控制，请使用 GetDocumentTextEx 方法。
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
                    // 获取正文内容（不包括页眉页脚）
                    // 表格中的文本会按单元格顺序提取
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
                // 注意：doc 是通过 app.ActiveDocument 获取的引用，不是本方法创建的，不应该释放
            }
        }

        /// <summary>
        /// 获取文档文本（高级版本）
        /// 遍历文档段落，更好地处理表格、列表等结构
        /// </summary>
        public static string GetDocumentTextEx(Word.Application app, bool includeTables = true, bool includeHeadersFooters = false)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            Word.Document doc = null;
            Word.Paragraphs paragraphs = null;
            Word.Range headerRange = null;
            Word.Range footerRange = null;
            var sb = new StringBuilder();

            try
            {
                doc = app.ActiveDocument;
                if (doc == null) throw new InvalidOperationException("当前未打开任何文档。");

                // 获取正文段落
                paragraphs = doc.Content.Paragraphs;
                int count = paragraphs.Count;

                for (int i = 1; i <= count; i++)
                {
                    Word.Paragraph para = null;
                    Word.Range paraRange = null;
                    
                    try
                    {
                        para = paragraphs[i];
                        paraRange = para.Range;
                        
                        string text = paraRange.Text?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(text))
                        {
                            // 检测是否在表格中
                            bool inTable = false;
                            try
                            {
                                inTable = (bool)paraRange.Information[Word.WdInformation.wdWithInTable];
                            }
                            catch { }

                            if (inTable && !includeTables)
                                continue; // 跳过表格内容

                            if (sb.Length > 0)
                                sb.AppendLine();
                            
                            sb.Append(text);
                        }
                    }
                    finally
                    {
                        if (paraRange != null) Marshal.ReleaseComObject(paraRange);
                        if (para != null) Marshal.ReleaseComObject(para);
                    }
                }

                // 可选：添加页眉页脚内容
                if (includeHeadersFooters)
                {
                    foreach (Word.Section section in doc.Sections)
                    {
                        Word.HeadersFooters headers = null;
                        Word.HeadersFooters footers = null;
                        
                        try
                        {
                            headers = section.Headers;
                            footers = section.Footers;

                            // 处理页眉
                            foreach (Word.HeaderFooter header in headers)
                            {
                                try
                                {
                                    headerRange = header.Range;
                                    string headerText = headerRange.Text?.Trim() ?? string.Empty;
                                    if (!string.IsNullOrEmpty(headerText))
                                    {
                                        sb.AppendLine();
                                        sb.AppendLine($"[页眉] {headerText}");
                                    }
                                }
                                finally
                                {
                                    if (headerRange != null) 
                                    {
                                        Marshal.ReleaseComObject(headerRange);
                                        headerRange = null;
                                    }
                                }
                            }

                            // 处理页脚
                            foreach (Word.HeaderFooter footer in footers)
                            {
                                try
                                {
                                    footerRange = footer.Range;
                                    string footerText = footerRange.Text?.Trim() ?? string.Empty;
                                    if (!string.IsNullOrEmpty(footerText))
                                    {
                                        sb.AppendLine();
                                        sb.AppendLine($"[页脚] {footerText}");
                                    }
                                }
                                finally
                                {
                                    if (footerRange != null) 
                                    {
                                        Marshal.ReleaseComObject(footerRange);
                                        footerRange = null;
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (headers != null) Marshal.ReleaseComObject(headers);
                            if (footers != null) Marshal.ReleaseComObject(footers);
                        }
                    }
                }

                string result = sb.ToString().Trim();
                if (string.IsNullOrWhiteSpace(result))
                    throw new InvalidOperationException("文档正文为空。");

                return result;
            }
            finally
            {
                if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                if (headerRange != null) Marshal.ReleaseComObject(headerRange);
                if (footerRange != null) Marshal.ReleaseComObject(footerRange);
                // 注意：doc 是引用，不应该释放
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
        /// 应用多个修订到文档（按位置倒序处理，避免偏移问题）
        /// </summary>
        public List<ProofreadIssueItem> ApplyRevisions(List<ProofreadIssueItem> items)
        {
            var processedItems = new List<ProofreadIssueItem>();
            if (!IsDocumentValid()) return processedItems;

            // 为每个项目找到匹配位置，然后按位置倒序排列（从后往前处理）
            var itemsWithPosition = new List<(ProofreadIssueItem item, int start, int end)>();
            
            // 第一步：为所有项目查找匹配位置（此时文档还未修改）
            foreach (var item in items)
            {
                if (item == null) continue;
                
                var (found, start, end) = FindTextPosition(item.Original);
                if (found)
                {
                    itemsWithPosition.Add((item, start, end));
                }
            }
            
            // 第二步：按位置倒序排列（从文档末尾到开头）
            // 这样替换时不会影响前面（文档头部）的位置
            itemsWithPosition = itemsWithPosition.OrderByDescending(x => x.start).ToList();
            
            // 第三步：逐个替换
            foreach (var (item, start, end) in itemsWithPosition)
            {
                try
                {
                    var commentBuilder = new StringBuilder();
                    commentBuilder.AppendLine($"【第{item.Index}处】类型：{item.Type}{(string.IsNullOrEmpty(item.Severity) ? "" : $"｜严重度：{item.Severity}")}");
                    commentBuilder.AppendLine($"原文：{item.Original}");
                    commentBuilder.AppendLine($"修改：{item.Modified}");
                    commentBuilder.AppendLine($"理由：{item.Reason}");

                    // 使用记录的位置直接替换，而不是重新 Find
                    if (ApplyRevisionAtRange(start, end, item.Original, item.Modified, commentBuilder.ToString(), out int newStart, out int newEnd))
                    {
                        item.DocumentStart = newStart;
                        item.DocumentEnd = newEnd;
                        processedItems.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyRevisions] 处理项目时出错: {ex.Message}");
                }
            }

            // 按原始顺序返回处理结果（保持索引顺序）
            return processedItems.OrderBy(i => i.Index).ToList();
        }
        
        /// <summary>
        /// 查找文本在文档中的位置
        /// 策略：1) 优先精确匹配（MatchCase=true, MatchWholeWord=true）
        ///       2) 如果失败，尝试宽松匹配（对长文本更严格，短文本更宽松）
        ///       3) 对短文本（<=2字）必须整词匹配，避免匹配到错误位置
        /// </summary>
        public (bool found, int start, int end) FindTextPosition(string text)
        {
            if (string.IsNullOrEmpty(text)) return (false, -1, -1);
            
            Word.Range searchRange = null;
            Word.Find find = null;
            
            try
            {
                searchRange = _document.Content;
                find = searchRange.Find;
                
                // 短文本（1-2个字符）必须整词匹配，避免 "的" 匹配到所有 "的" 字
                bool isShortText = text.Length <= 2;
                
                // 第1步：尝试精确匹配（区分大小写 + 整词匹配）
                bool found = find.Execute(FindText: text, 
                                          MatchCase: true, 
                                          MatchWholeWord: true,
                                          MatchWildcards: false, 
                                          Forward: true, 
                                          Wrap: Word.WdFindWrap.wdFindStop);
                
                if (found)
                {
                    // 验证：短文本直接返回，长文本验证内容
                    if (!isShortText || searchRange.Text == text)
                    {
                        return (true, searchRange.Start, searchRange.End);
                    }
                }
                
                // 第2步：如果精确匹配失败，尝试不区分大小写但仍整词匹配
                find = searchRange.Find; // 重新获取 Find 对象
                found = find.Execute(FindText: text, 
                                     MatchCase: false, 
                                     MatchWholeWord: true,
                                     MatchWildcards: false, 
                                     Forward: true, 
                                     Wrap: Word.WdFindWrap.wdFindStop);
                
                if (found)
                {
                    return (true, searchRange.Start, searchRange.End);
                }
                
                // 第3步：对于长文本（>5字），允许宽松匹配（不整词），因为可能是句子片段
                if (text.Length > 5)
                {
                    find = searchRange.Find;
                    found = find.Execute(FindText: text, 
                                         MatchCase: false, 
                                         MatchWholeWord: false,
                                         MatchWildcards: false, 
                                         Forward: true, 
                                         Wrap: Word.WdFindWrap.wdFindStop);
                    
                    if (found)
                    {
                        return (true, searchRange.Start, searchRange.End);
                    }
                }
            }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
            
            return (false, -1, -1);
        }
        
        /// <summary>
        /// 在指定范围内应用修订（使用确切位置）
        /// </summary>
        public bool ApplyRevisionAtRange(int start, int end, string original, string modified, string commentText, out int newStart, out int newEnd)
        {
            newStart = -1;
            newEnd = -1;
            
            if (!IsDocumentValid()) return false;
            if (start < 0 || end <= start) return false;
            
            Word.Range range = null;
            Word.Comment comment = null;
            
            try
            {
                // 直接定位到指定范围
                range = _document.Range(start, end);
                
                // 验证内容是否匹配（可选的安全检查）
                if (range.Text != original)
                {
                    // 内容不匹配，尝试用 Find 在当前范围附近查找
                    // 使用整词匹配提高精度，避免短文本匹配错误
                    Word.Find find = range.Find;
                    bool isShortText = original?.Length <= 2;
                    
                    // 优先使用整词匹配
                    bool found = find.Execute(FindText: original, 
                                              MatchCase: false, 
                                              MatchWholeWord: true,
                                              MatchWildcards: false, 
                                              Forward: true, 
                                              Wrap: Word.WdFindWrap.wdFindStop);
                    
                    // 短文本必须整词匹配，长文本可以放宽
                    if (!found && !isShortText && original?.Length > 5)
                    {
                        find = range.Find;
                        found = find.Execute(FindText: original, 
                                             MatchCase: false, 
                                             MatchWholeWord: false,
                                             MatchWildcards: false, 
                                             Forward: true, 
                                             Wrap: Word.WdFindWrap.wdFindStop);
                    }
                    
                    Marshal.ReleaseComObject(find);
                    
                    if (!found) return false;
                }
                
                bool oldTrackRevisions = _document.TrackRevisions;
                try
                {
                    _document.TrackRevisions = true;
                    range.Text = modified;
                    comment = _document.Comments.Add(range, commentText);
                    newStart = range.Start;
                    newEnd = range.End;
                    return true;
                }
                finally
                {
                    _document.TrackRevisions = oldTrackRevisions;
                }
            }
            finally
            {
                if (comment != null) Marshal.ReleaseComObject(comment);
                if (range != null) Marshal.ReleaseComObject(range);
            }
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
        /// 策略：优先使用原文搜索（最可靠），搜索失败时尝试使用缓存位置作为备用
        /// 注意：DocumentStart/End 在多修订后可能偏移，仅作为缓存加速使用
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

            // 方案1：优先使用原文搜索（最可靠，不受修订偏移影响）
            if (NavigateBySearch(item.Original))
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToIssue] 通过原文搜索定位成功: '{item.Original.Substring(0, Math.Min(20, item.Original.Length))}...'");
                return;
            }

            // 方案2：原文搜索失败，尝试使用缓存位置（可能已偏移，仅作为备用）
            if (item.DocumentStart >= 0 && item.DocumentEnd > item.DocumentStart)
            {
                System.Diagnostics.Debug.WriteLine($"[NavigateToIssue] 原文搜索失败，尝试使用缓存位置: {item.DocumentStart}-{item.DocumentEnd}");
                if (NavigateToRange(item.DocumentStart, item.DocumentEnd))
                    return;
            }

            // 都失败了
            throw new InvalidOperationException("无法在文档中找到该位置，可能文本已被修改或删除。");
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
                    // 注意：doc 是通过 app.ActiveDocument 获取的引用，不是本方法创建的，不应该释放
                    errorMessage = "文档已被释放，请重新打开";
                    return false;
                }

                service = new WordDocumentService(app, doc);
                return true;
            }
            catch (Exception ex)
            {
                // 注意：doc 是通过 app.ActiveDocument 获取的引用，不是本方法创建的，不应该释放
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
