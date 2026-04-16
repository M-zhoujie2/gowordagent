/**
 * WPS 文档操作服务
 * 封装 WPS JS API
 */

var DocumentService = {
    /**
     * 获取文档文本（带偏移量信息）
     * @returns {Array} 段落数组，包含索引、文本、偏移量
     */
    getDocumentText: function() {
        try {
            var app = wps.WpsApplication();
            if (!app) {
                throw new Error('无法获取 WPS 应用程序');
            }
            
            var doc = app.ActiveDocument;
            
            if (!doc) {
                throw new Error('没有打开的文档，请先打开一个文档');
            }
            
            // 检查文档是否已保存（避免处理临时文档时的问题）
            try {
                var _ = doc.Content.Text;
            } catch (e) {
                throw new Error('文档状态异常，请保存后重试');
            }
            
            var paragraphs = [];
            var offset = 0;
            var paraCount = doc.Paragraphs.Count;
            
            // 限制最大段落数（防止超大文档导致性能问题）
            var maxParagraphs = 10000;
            if (paraCount > maxParagraphs) {
                console.warn('文档段落过多 (' + paraCount + ')，只处理前 ' + maxParagraphs + ' 段');
                paraCount = maxParagraphs;
            }
            
            for (var i = 1; i <= paraCount; i++) {
                try {
                    var p = doc.Paragraphs.Item(i);
                    var text = p.Range.Text || '';
                    
                    // 过滤空段落
                    if (text.trim().length === 0) {
                        offset += text.length;
                        continue;
                    }
                    
                    paragraphs.push({
                        index: i - 1,  // 0-based index
                        start: offset,
                        end: offset + text.length,
                        text: text
                    });
                    
                    offset += text.length;
                } catch (paraError) {
                    console.error('读取第 ' + i + ' 段时出错:', paraError);
                    // 继续处理下一段
                }
            }
            
            if (paragraphs.length === 0) {
                throw new Error('文档内容为空或无法读取');
            }
            
            console.log('DocumentService: 成功提取 ' + paragraphs.length + ' 个段落');
            return paragraphs;
            
        } catch (e) {
            console.error('DocumentService.getDocumentText 错误:', e);
            throw e;
        }
    },

    /**
     * 在指定偏移量位置应用修订
     * @param {number} startOffset - 起始偏移量
     * @param {number} endOffset - 结束偏移量
     * @param {string} replacement - 替换文本
     * @param {string} comment - 批注内容
     * @returns {boolean} 是否成功
     */
    applyAtOffset: function(startOffset, endOffset, replacement, comment) {
        try {
            // 参数验证
            if (typeof startOffset !== 'number' || typeof endOffset !== 'number') {
                console.error('applyAtOffset: 偏移量必须是数字');
                return false;
            }
            
            if (startOffset < 0 || endOffset <= startOffset) {
                console.error('applyAtOffset: 偏移量范围无效', startOffset, endOffset);
                return false;
            }
            
            if (!replacement && !comment) {
                console.warn('applyAtOffset: 替换文本和批注都为空');
                return false;
            }
            
            var app = wps.WpsApplication();
            if (!app) {
                console.error('applyAtOffset: 无法获取 WPS 应用程序');
                return false;
            }
            
            var doc = app.ActiveDocument;
            if (!doc) {
                console.error('applyAtOffset: 没有活动文档');
                return false;
            }
            
            // 获取文档长度并验证偏移量
            var docLength = doc.Content.Text.length;
            if (endOffset > docLength) {
                console.warn('applyAtOffset: 结束偏移量超出文档长度，自动调整', 
                    endOffset, '->', docLength);
                endOffset = docLength;
            }
            
            if (startOffset >= docLength) {
                console.error('applyAtOffset: 起始偏移量超出文档长度');
                return false;
            }
            
            // 定位到指定范围
            var range;
            try {
                range = doc.Range(startOffset, endOffset);
            } catch (e) {
                console.error('applyAtOffset: 创建 Range 失败', e);
                return false;
            }
            
            if (!range) {
                console.error('applyAtOffset: Range 为空');
                return false;
            }
            
            // 保存原始修订状态
            var oldTrackRevisions = false;
            try {
                oldTrackRevisions = doc.TrackRevisions;
            } catch (e) {
                console.warn('applyAtOffset: 无法获取 TrackRevisions 状态');
            }
            
            // 开启修订模式
            try {
                doc.TrackRevisions = true;
            } catch (e) {
                console.warn('applyAtOffset: 无法开启修订模式', e);
                // 继续执行，只是不会有修订标记
            }
            
            try {
                // 删除原文
                try {
                    range.Delete();
                } catch (e) {
                    console.error('applyAtOffset: 删除原文失败', e);
                    return false;
                }
                
                // 插入新文本
                if (replacement) {
                    try {
                        range.InsertAfter(replacement);
                    } catch (e) {
                        console.error('applyAtOffset: 插入新文本失败', e);
                        return false;
                    }
                }
                
                // 添加批注
                if (comment) {
                    try {
                        doc.Comments.Add(range, comment);
                    } catch (e) {
                        console.warn('applyAtOffset: 添加批注失败', e);
                        // 批注失败不视为整体失败
                    }
                }
                
                return true;
                
            } finally {
                // 恢复原始修订状态
                try {
                    doc.TrackRevisions = oldTrackRevisions;
                } catch (e) {
                    console.warn('applyAtOffset: 恢复 TrackRevisions 失败', e);
                }
            }
            
        } catch (e) {
            console.error('applyAtOffset 错误:', e);
            return false;
        }
    },

    /**
     * 导航到指定偏移量位置
     * @param {number} startOffset - 起始偏移量
     * @param {number} endOffset - 结束偏移量
     * @returns {boolean} 是否成功
     */
    navigateToOffset: function(startOffset, endOffset) {
        try {
            // 参数验证
            if (typeof startOffset !== 'number' || typeof endOffset !== 'number') {
                console.error('navigateToOffset: 偏移量必须是数字');
                return false;
            }
            
            if (startOffset < 0 || endOffset <= startOffset) {
                console.error('navigateToOffset: 偏移量范围无效');
                return false;
            }
            
            var app = wps.WpsApplication();
            if (!app) {
                console.error('navigateToOffset: 无法获取 WPS 应用程序');
                return false;
            }
            
            var doc = app.ActiveDocument;
            if (!doc) {
                console.error('navigateToOffset: 没有活动文档');
                return false;
            }
            
            // 验证偏移量范围
            var docLength = doc.Content.Text.length;
            if (startOffset >= docLength) {
                console.error('navigateToOffset: 起始偏移量超出文档长度');
                return false;
            }
            
            if (endOffset > docLength) {
                endOffset = docLength;
            }
            
            // 定位并选中
            var range;
            try {
                range = doc.Range(startOffset, endOffset);
            } catch (e) {
                console.error('navigateToOffset: 创建 Range 失败', e);
                return false;
            }
            
            if (!range) {
                console.error('navigateToOffset: Range 为空');
                return false;
            }
            
            // 选中范围
            try {
                range.Select();
            } catch (e) {
                console.error('navigateToOffset: Select 失败', e);
                return false;
            }
            
            // 滚动到视图中
            var window = app.ActiveWindow;
            if (window) {
                try {
                    if (window.ScrollIntoView) {
                        window.ScrollIntoView(range);
                    } else {
                        // 备用方案：使用 Range 的 ScrollIntoView
                        range.ScrollIntoView();
                    }
                } catch (e) {
                    console.warn('navigateToOffset: ScrollIntoView 失败', e);
                    // 滚动失败不算整体失败
                }
            }
            
            return true;
            
        } catch (e) {
            console.error('navigateToOffset 错误:', e);
            return false;
        }
    },

    /**
     * 检查文档是否有效
     * @returns {boolean}
     */
    isDocumentValid: function() {
        try {
            var app = wps.WpsApplication();
            if (!app) return false;
            
            var doc = app.ActiveDocument;
            if (!doc) return false;
            
            // 尝试访问文档属性
            var _ = doc.Content.Text;
            return true;
            
        } catch (e) {
            return false;
        }
    }
};
