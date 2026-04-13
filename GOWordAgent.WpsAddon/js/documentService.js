/**
 * WPS 文档操作服务
 * 封装 WPS JS API
 */

var DocumentService = {
    /**
     * 获取文档文本（带偏移量信息）
     */
    getDocumentText: function() {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        if (!doc) {
            throw new Error('没有打开的文档');
        }
        
        var paragraphs = [];
        var offset = 0;
        
        for (var i = 1; i <= doc.Paragraphs.Count; i++) {
            var p = doc.Paragraphs.Item(i);
            var text = p.Range.Text;
            
            paragraphs.push({
                index: i - 1,
                start: offset,
                end: offset + text.length,
                text: text
            });
            
            offset += text.length;
        }
        
        return paragraphs;
    },

    /**
     * 在指定偏移量位置应用修订
     */
    applyAtOffset: function(startOffset, endOffset, replacement, comment) {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        // 定位到指定范围
        var range = doc.Range(startOffset, endOffset);
        
        // 开启修订模式
        var oldTrackRevisions = doc.TrackRevisions;
        doc.TrackRevisions = true;
        
        try {
            // 删除原文并插入新文本
            range.Delete();
            range.InsertAfter(replacement);
            
            // 添加批注
            if (comment) {
                doc.Comments.Add(range, comment);
            }
        } finally {
            doc.TrackRevisions = oldTrackRevisions;
        }
    },

    /**
     * 导航到指定偏移量位置
     */
    navigateToOffset: function(startOffset, endOffset) {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        
        var range = doc.Range(startOffset, endOffset);
        range.Select();
        
        // 滚动到视图中
        var window = app.ActiveWindow;
        if (window && window.ScrollIntoView) {
            window.ScrollIntoView(range);
        }
    },

    /**
     * 获取选中的文本
     */
    getSelectedText: function() {
        var app = wps.WpsApplication();
        var selection = app.Selection;
        return selection ? selection.Text : '';
    },

    /**
     * 获取文档总字符数
     */
    getDocumentLength: function() {
        var app = wps.WpsApplication();
        var doc = app.ActiveDocument;
        return doc ? doc.Content.Text.length : 0;
    }
};
