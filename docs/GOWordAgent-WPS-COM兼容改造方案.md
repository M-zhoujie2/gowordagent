# GOWordAgent WPS COM 兼容改造方案（Windows 环境）

> 版本：v1.0 | 日期：2026-03-31 | 基于源码分析
> 适用环境：Windows + WPS Pro（个人版/专业版）

---

## 一、方案概述

### 核心思路

WPS Windows 版**暴露了与 Word 兼容的 COM 接口**，ProgID 为 `KWps.Application`。本方案通过以下策略让现有 VSTO 插件（最小改动）运行在 WPS 上：

1. **不换架构、不拆前后端**，保持 VSTO 单体结构
2. 抽象 Word COM 调用层，根据宿主自动切换 Word/WPS 实现
3. 替换 VSTO 的加载机制为 COM Add-in（WPS 支持标准 COM Add-in）

### 可行性基础

WPS Windows 版支持的 COM 能力：

| 能力 | WPS 支持情况 | 本项目是否依赖 |
|------|------------|--------------|
| `Application` 对象 | ✅ `KWps.Application` | ✅ 是 |
| `ActiveDocument` | ✅ | ✅ 是 |
| `Range` 文本操作 | ✅ 基础操作 | ✅ 是 |
| `Range.Find` | ✅ | ✅ 是 |
| `TrackRevisions` | ✅ | ✅ 是 |
| `Comments.Add` | ✅ | ✅ 是 |
| `Paragraphs` | ✅ | ✅ 是 |
| `Selection` | ✅ | ✅ 是 |
| `ScrollIntoView` | ⚠️ 需实测 | ⚠️ 有降级 |
| `CustomTaskPane` | ❌ **不支持** | ✅ 是 |
| VSTO Runtime | ❌ **不支持** | ✅ 是 |
| `WdFindWrap` 等枚举 | ⚠️ 部分 | ⚠️ 有替代 |

**关键障碍**：WPS 不支持 VSTO 的 `CustomTaskPane` 和 `Ribbon` 扩展。

---

## 二、架构改动

### 2.1 改动前后对比

```
改动前（Word 专用）：
  VSTO 插件
  ├── ThisAddIn (VSTO 入口)
  ├── GOWordAgentRibbon (VSTO Ribbon)
  ├── GOWordAgentPaneHost (WinForms → WPF 宿主)
  ├── GOWordAgentPaneWpf.xaml.cs (WPF 侧边栏 UI)
  ├── WordDocumentService.cs (Word COM)
  ├── WordProofreadController.cs (调用 WordDocumentService)
  └── Services/ (LLM、缓存、配置)

改动后（Word + WPS 双宿主）：
  COM Add-in（注册表加载，VSTO 和 WPS 都支持）
  ├── ThisAddIn (改为标准 COM Add-in 入口)
  ├── UI 层
  │   ├── Word 版：保持 VSTO Ribbon + CustomTaskPane + WPF
  │   └── WPS 版：独立 WPF Window（浮动窗口替代侧边栏）
  ├── WordDocumentService.cs
  │   ├── IWordInterop.cs（新增接口）
  │   ├── MsWordInterop.cs（Word 实现，原逻辑）
  │   └── WpsWordInterop.cs（WPS 实现，COM 差异适配）
  ├── WordProofreadController.cs（改动最小）
  └── Services/（基本不改）
```

### 2.2 关键设计决策

| 决策点 | 方案 | 理由 |
|--------|------|------|
| 插件加载方式 | 标准COM Add-in | WPS 不支持 VSTO 但支持 COM Add-in（`IDTExtensibility2`） |
| UI 宿主 | Word用CustomTaskPane，WPS用WPF Window | WPS 无 CustomTaskPane，浮动窗口是唯一选择 |
| COM 适配 | 接口抽象 + 两个实现 | 隔离差异，切换成本低 |
| 枚举常量 | WPS用数值替代Word枚举 | WPS COM 不注册 Word 枚举类型 |

---

## 三、逐文件改动方案

### 3.1 新增文件

#### 3.1.1 IWordInterop.cs（COM 操作抽象接口）

```csharp
using System;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// Word/WPS 文档操作抽象接口
    /// 提取自 WordDocumentService.cs 的公开方法
    /// </summary>
    public interface IWordInterop : IDisposable
    {
        /// <summary>获取当前活动文档全文或选中文本</summary>
        string GetDocumentText();

        /// <summary>检查文档是否有效</summary>
        bool IsDocumentValid();

        /// <summary>查找文本位置（三级匹配）</summary>
        (bool found, int start, int end) FindTextPosition(string text);

        /// <summary>在指定范围应用修订</summary>
        bool ApplyRevisionAtRange(int start, int end, string original, 
            string modified, string commentText, out int newStart, out int newEnd);

        /// <summary>导航到指定范围</summary>
        bool NavigateToRange(int start, int end);

        /// <summary>通过文本搜索导航</summary>
        bool NavigateBySearch(string originalText);

        /// <summary>激活窗口</summary>
        bool ActivateWindow();

        /// <summary>获取文档名称</summary>
        string GetDocumentName();
    }
}
```

#### 3.1.2 HostType.cs（宿主类型）

```csharp
namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// 当前运行宿主
    /// </summary>
    public enum HostType
    {
        Word,
        Wps
    }
}
```

#### 3.1.3 HostDetector.cs（宿主检测）

```csharp
using System;
using System.Runtime.InteropServices;

namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// 检测当前运行的宿主是 Word 还是 WPS
    /// 策略：检查进程名
    /// </summary>
    public static class HostDetector
    {
        /// <summary>
        /// 在插件加载时检测宿主类型
        /// </summary>
        public static HostType Detect()
        {
            string processName = System.Diagnostics.Process.GetCurrentProcess().ProcessName.ToLower();
            if (processName.Contains("wps") || processName.Contains("kwps"))
                return HostType.Wps;
            if (processName.Contains("winword"))
                return HostType.Word;
            // 默认按 Word 处理
            return HostType.Word;
        }
    }
}
```

### 3.2 改造现有文件

#### 3.2.1 WordDocumentService.cs → 拆分为接口 + 两个实现

**原文件保留**，改名为 `MsWordInterop.cs`，实现 `IWordInterop`：

```csharp
// MsWordInterop.cs — 原 WordDocumentService.cs 改造
using System;
using System.Runtime.InteropServices;
using System.Text;
using GOWordAgentAddIn.Interop;
using GOWordAgentAddIn.Models;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// MS Word COM 实现
    /// 从原 WordDocumentService.cs 提取，逻辑基本不变
    /// </summary>
    public class MsWordInterop : IWordInterop
    {
        private readonly Word.Application _application;
        private readonly Word.Document _document;
        private bool _disposed;

        public MsWordInterop(Word.Application application, Word.Document document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        // === 以下方法从原 WordDocumentService.cs 迁移，逻辑不变 ===

        public string GetDocumentText()
        {
            return GetDocumentTextStatic(_application);
        }

        public bool IsDocumentValid()
        {
            Word.Range content = null;
            try
            {
                content = _document.Content;
                var _ = content.Text;
                return true;
            }
            catch (COMException) { return false; }
            finally { if (content != null) Marshal.ReleaseComObject(content); }
        }

        public (bool found, int start, int end) FindTextPosition(string text)
        {
            // ... 原逻辑完全不变 ...
        }

        public bool ApplyRevisionAtRange(int start, int end, string original,
            string modified, string commentText, out int newStart, out int newEnd)
        {
            // ... 原逻辑完全不变 ...
        }

        public bool NavigateToRange(int start, int end)
        {
            // ... 原逻辑不变，ScrollIntoView 保留 try-catch 降级 ...
        }

        public bool NavigateBySearch(string originalText)
        {
            // ... 原逻辑不变 ...
        }

        public bool ActivateWindow()
        {
            try { _application.Activate(); return true; }
            catch (COMException) { return false; }
        }

        public string GetDocumentName()
        {
            try { return _document.Name; }
            catch { return "未知文档"; }
        }

        // 内部方法保持私有：GetDocumentTextStatic, TryApplyWithRevisions,
        // ApplyDegradedRevision 等全部保留不变

        // Dispose 保持不变

        // ====== 新增：静态工厂方法（从原 WordDocumentServiceFactory 迁移） ======
        public static bool TryCreate(Word.Application app, Word.Document doc,
            out IWordInterop interop, out string errorMessage)
        {
            interop = null;
            errorMessage = null;
            try
            {
                if (app == null) { errorMessage = "无法访问 Word 应用"; return false; }
                if (doc == null) { errorMessage = "文档为空"; return false; }
                try { var _ = doc.Content.Text; }
                catch (COMException) { errorMessage = "文档已被释放"; return false; }

                interop = new MsWordInterop(app, doc);
                return true;
            }
            catch (Exception ex) { errorMessage = ex.Message; return false; }
        }
    }
}
```

#### 3.2.2 新增 WpsWordInterop.cs（WPS COM 适配层）

```csharp
// WpsWordInterop.cs — WPS COM 适配
using System;
using System.Runtime.InteropServices;
using System.Text;
using GOWordAgentAddIn.Interop;
using GOWordAgentAddIn.Models;

namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// WPS COM 实现
    /// 使用 late binding（dynamic）避免直接引用 WPS 类型库
    /// 
    /// ⚠️ 关键差异：
    /// 1. WPS COM 不注册 Word 枚举（WdFindWrap 等），需用数值
    /// 2. 部分 API 行为可能有细微差异
    /// 3. ScrollIntoView 可能不支持
    /// 4. TrackRevisions 行为需实测
    /// </summary>
    public class WpsWordInterop : IWordInterop
    {
        private readonly dynamic _application;  // KWps.Application
        private readonly dynamic _document;      // Document
        private bool _disposed;

        public WpsWordInterop(dynamic application, dynamic document)
        {
            _application = application;
            _document = document;
        }

        public string GetDocumentText()
        {
            try
            {
                // 尝试获取选中文本
                dynamic selection = null;
                try
                {
                    selection = _application.ActiveWindow.Selection;
                    string selectedText = selection.Text;
                    if (!string.IsNullOrWhiteSpace(selectedText))
                        return selectedText;
                }
                catch { }
                finally
                {
                    if (selection != null) Marshal.ReleaseComObject(selection);
                }

                // 获取全文
                dynamic content = _document.Content;
                string text = content.Text;
                Marshal.ReleaseComObject(content);

                if (string.IsNullOrWhiteSpace(text))
                    throw new InvalidOperationException("文档正文为空。");
                return text;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"获取文档失败: {ex.Message}");
            }
        }

        public bool IsDocumentValid()
        {
            try
            {
                dynamic content = _document.Content;
                var _ = content.Text;
                Marshal.ReleaseComObject(content);
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// WPS Find.Execute — 用数值替代 Word 枚举
        /// 
        /// Word: Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, 
        ///        Forward, Wrap, Format, ReplaceWith, Replace)
        /// WPS:  相同参数顺序，但 Wrap 枚举用数值
        ///   wdFindStop = 0
        ///   wdFindContinue = 1
        ///   wdFindAsk = 2
        /// </summary>
        public (bool found, int start, int end) FindTextPosition(string text)
        {
            if (string.IsNullOrEmpty(text)) return (false, -1, -1);

            dynamic searchRange = null;
            try
            {
                searchRange = _document.Range(0, 0); // 从文档开头
                dynamic find = searchRange.Find;

                // 第 1 步：精确匹配
                find.ClearFormatting();
                bool found = (bool)find.Execute(
                    FindText: text,
                    MatchCase: true,
                    MatchWholeWord: true,
                    MatchWildcards: false,
                    Forward: true,
                    Wrap: 0  // wdFindStop = 0，用数值替代枚举
                );
                if (found)
                    return (true, (int)searchRange.Start, (int)searchRange.End);

                // 第 2 步：不区分大小写 + 整词
                found = (bool)find.Execute(
                    FindText: text, MatchCase: false, MatchWholeWord: true,
                    MatchWildcards: false, Forward: true, Wrap: 0
                );
                if (found)
                    return (true, (int)searchRange.Start, (int)searchRange.End);

                // 第 3 步：长文本宽松匹配
                if (text.Length > 5)
                {
                    found = (bool)find.Execute(
                        FindText: text, MatchCase: false, MatchWholeWord: false,
                        MatchWildcards: false, Forward: true, Wrap: 0
                    );
                    if (found)
                        return (true, (int)searchRange.Start, (int)searchRange.End);
                }

                return (false, -1, -1);
            }
            finally
            {
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        public bool ApplyRevisionAtRange(int start, int end, string original,
            string modified, string commentText, out int newStart, out int newEnd)
        {
            newStart = -1;
            newEnd = -1;

            if (!IsDocumentValid() || start < 0 || end <= start)
                return false;

            dynamic range = null;
            dynamic comment = null;
            try
            {
                range = _document.Range(start, end);

                // 验证内容匹配
                if ((string)range.Text != original)
                {
                    // 尝试 Find 定位
                    dynamic find = range.Find;
                    bool found = (bool)find.Execute(original, false, true, false, true, 0);
                    if (!found && original.Length > 5)
                    {
                        Marshal.ReleaseComObject(find);
                        find = range.Find;
                        found = (bool)find.Execute(original, false, false, false, true, 0);
                    }
                    if (!found) return false;
                }

                // 尝试修订模式
                bool oldTrack = (bool)_document.TrackRevisions;
                try
                {
                    _document.TrackRevisions = true;
                    range.Text = modified;

                    // 添加批注
                    try
                    {
                        // ⚠️ WPS Comments.Add 参数顺序可能与 Word 不同
                        // Word: Comments.Add(Range, Text)
                        // WPS:  可能相同，也可能反过来
                        comment = _document.Comments.Add(range, commentText);
                    }
                    catch (COMException ex)
                    {
                        // 尝试另一种参数顺序
                        try
                        {
                            comment = _document.Comments.Add(commentText, range);
                        }
                        catch (COMException ex2)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[WpsWordInterop] 批注添加失败: {ex2.Message}");
                        }
                    }

                    newStart = (int)range.Start;
                    newEnd = (int)range.End;
                    return true;
                }
                catch (COMException ex)
                {
                    // 降级：直接替换
                    System.Diagnostics.Debug.WriteLine(
                        $"[WpsWordInterop] 修订失败，降级处理: {ex.Message}");
                    range.Text = $"[原文：{original}] {modified}";
                    try { comment = _document.Comments.Add(range, commentText + "\n⚠️ 修订不可用"); }
                    catch { }
                    newStart = (int)range.Start;
                    newEnd = (int)range.End;
                    return true;
                }
                finally
                {
                    try { _document.TrackRevisions = oldTrack; }
                    catch { }
                }
            }
            finally
            {
                if (comment != null) Marshal.ReleaseComObject(comment);
                if (range != null) Marshal.ReleaseComObject(range);
            }
        }

        public bool NavigateToRange(int start, int end)
        {
            if (!IsDocumentValid() || start < 0 || end <= start)
                return false;

            try
            {
                _application.Activate();

                dynamic range = _document.Range(start, end);
                range.Select();

                // ⚠️ WPS ScrollIntoView 可能不支持
                try
                {
                    _application.ActiveWindow.ScrollIntoView(range);
                }
                catch (COMException)
                {
                    // 降级：仅选中，不滚动
                    System.Diagnostics.Debug.WriteLine("[WpsWordInterop] ScrollIntoView 不支持");
                }
                finally
                {
                    Marshal.ReleaseComObject(range);
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WpsWordInterop] 导航失败: {ex.Message}");
                return false;
            }
        }

        public bool NavigateBySearch(string originalText)
        {
            if (!IsDocumentValid() || string.IsNullOrWhiteSpace(originalText))
                return false;

            var (found, start, end) = FindTextPosition(originalText);
            if (found) return NavigateToRange(start, end);
            return false;
        }

        public bool ActivateWindow()
        {
            try { _application.Activate(); return true; }
            catch { return false; }
        }

        public string GetDocumentName()
        {
            try { return (string)_document.Name; }
            catch { return "未知文档"; }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                // _application 和 _document 是外部引用，不释放
                _disposed = true;
            }
        }

        // ====== 工厂方法 ======
        public static bool TryCreate(dynamic application, dynamic document,
            out IWordInterop interop, out string errorMessage)
        {
            interop = null;
            errorMessage = null;
            try
            {
                if (application == null) { errorMessage = "无法访问 WPS 应用"; return false; }
                if (document == null) { errorMessage = "文档为空"; return false; }
                try { var _ = (string)document.Content.Text; }
                catch { errorMessage = "文档已被释放"; return false; }

                interop = new WpsWordInterop(application, document);
                return true;
            }
            catch (Exception ex) { errorMessage = ex.Message; return false; }
        }
    }
}
```

#### 3.2.3 InteropFactory.cs（统一工厂）

```csharp
using System;
using System.Runtime.InteropServices;
using GOWordAgentAddIn.Interop;
using Word = Microsoft.Office.Interop.Word;

namespace GOWordAgentAddIn.Interop
{
    /// <summary>
    /// 文档操作工厂 — 根据宿主自动创建对应的 Interop 实现
    /// </summary>
    public static class InteropFactory
    {
        private static HostType _hostType;

        public static HostType CurrentHost => _hostType;

        /// <summary>
        /// 初始化（插件启动时调用一次）
        /// </summary>
        public static void Initialize(HostType hostType)
        {
            _hostType = hostType;
            System.Diagnostics.Debug.WriteLine($"[InteropFactory] 宿主: {_hostType}");
        }

        /// <summary>
        /// 为当前活动文档创建 Interop
        /// </summary>
        public static bool TryCreateForActiveDocument(out IWordInterop interop, out string errorMessage)
        {
            interop = null;
            errorMessage = null;

            try
            {
                if (_hostType == HostType.Word)
                {
                    // Word 路径：通过 VSTO Globals 获取
                    var app = Globals.ThisAddIn?.Application;
                    if (app == null) { errorMessage = "无法访问 Word"; return false; }
                    var doc = app.ActiveDocument;
                    if (doc == null) { errorMessage = "无活动文档"; return false; }
                    return MsWordInterop.TryCreate(app, doc, out interop, out errorMessage);
                }
                else
                {
                    // WPS 路径：通过 WPS COM 获取
                    return TryCreateForWps(out interop, out errorMessage);
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool TryCreateForWps(out IWordInterop interop, out string errorMessage)
        {
            dynamic app = null;
            dynamic doc = null;
            try
            {
                // 获取 WPS Application
                Type wpsType = Type.GetTypeFromProgID("KWps.Application");
                if (wpsType == null) wpsType = Type.GetTypeFromProgID("WPS.Application");
                if (wpsType == null) { errorMessage = "未找到 WPS"; return false; }

                // 方式 1：如果已有 WPS 实例（推荐）
                try
                {
                    app = Marshal.GetActiveObject("KWps.Application");
                }
                catch
                {
                    try { app = Marshal.GetActiveObject("WPS.Application"); }
                    catch { errorMessage = "WPS 未运行"; return false; }
                }

                doc = app.ActiveDocument;
                if (doc == null) { errorMessage = "WPS 无活动文档"; return false; }

                return WpsWordInterop.TryCreate(app, doc, out interop, out errorMessage);
            }
            catch (COMException ex)
            {
                errorMessage = $"WPS COM 错误: {ex.Message}";
                return false;
            }
            finally
            {
                // 不释放 app/doc，是引用
            }
        }
    }
}
```

#### 3.2.4 WordProofreadController.cs 改造

改动很小，把 `WordDocumentService` 替换为 `IWordInterop`：

```csharp
// WordProofreadController.cs — 改动清单

// 1. 新增 using
using GOWordAgentAddIn.Interop;

// 2. 替换字段类型
// 改前：
//     private WordDocumentService _documentService;
// 改后：
    private IWordInterop _documentService;  // ← 唯一的类型改动

// 3. TryGetDocumentService 方法改动
private bool TryGetDocumentService(out IWordInterop service, out string errorMessage)
{
    lock (_lock)
    {
        // 改前：检查绑定文档 + WordDocumentServiceFactory
        // 改后：用 InteropFactory 统一创建

        if (_documentService == null || !_documentService.IsDocumentValid())
        {
            _documentService?.Dispose();
            _documentService = null;

            // 改前：WordDocumentServiceFactory.TryCreateForDocument(app, doc, ...)
            // 改后：
            if (!InteropFactory.TryCreateForActiveDocument(out _documentService, out errorMessage))
            {
                service = null;
                return false;
            }
        }

        service = _documentService;
        errorMessage = null;
        return true;
    }
}

// 4. GetDocumentText 改动
public string GetDocumentText()
{
    try
    {
        // 改前：WordDocumentService.GetDocumentText(app)
        // 改后：
        if (!InteropFactory.TryCreateForActiveDocument(out var interop, out var err))
        {
            MessageBox.Show($"获取文档失败: {err}", "错误");
            return null;
        }
        string text = interop.GetDocumentText();
        return text;
    }
    catch (Exception ex)
    {
        MessageBox.Show($"获取文档失败: {ex.Message}", "错误");
        return null;
    }
}

// 5. 其余方法签名不变，内部调用 _documentService 的方法
//    因为 IWordInterop 接口与原 WordDocumentService 公开方法一致，
//    FindItemPositions、ApplyRevisions 等方法无需改动
```

**WordProofreadController.cs 总改动量：约 20 行**，核心是把 `WordDocumentService` 替换为 `IWordInterop`。

#### 3.2.5 ThisAddIn.cs 改造

```csharp
// ThisAddIn.cs — 改动清单

// 1. 新增字段
private HostType _hostType;

// 2. ThisAddIn_Startup 改造
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    Current = this;

    // 检测宿主
    _hostType = HostDetector.Detect();
    InteropFactory.Initialize(_hostType);

    System.Diagnostics.Debug.WriteLine($"[ThisAddIn] 宿主: {_hostType}");

    if (_hostType == HostType.Wps)
    {
        // WPS 模式：不初始化 CustomTaskPane（WPS 不支持）
        // 侧边栏通过浮动 WPF Window 实现
        InitializeWpsMode();
    }
    // Word 模式：保持原有延迟初始化逻辑不变
}

// 3. WPS 模式初始化
private WpfFloatingWindow _wpsWindow;

private void InitializeWpsMode()
{
    // 创建浮动 WPF 窗口（替代 CustomTaskPane）
    _wpsWindow = new WpfFloatingWindow();
    _wpsWindow.Closed += (s, e) => { _wpsWindow = null; };
}

/// <summary>
/// 获取或初始化面板（兼容 Word/WPS）
/// </summary>
public object GetOrInitializePane()
{
    if (_hostType == HostType.Wps)
    {
        // WPS: 显示/隐藏浮动窗口
        if (_wpsWindow == null)
            InitializeWpsMode();
        _wpsWindow.Visible = !_wpsWindow.Visible;
        return _wpsWindow;
    }
    else
    {
        // Word: 原有 CustomTaskPane 逻辑
        if (!_isPaneInitialized)
        {
            lock (_initLock)
            {
                if (!_isPaneInitialized)
                {
                    InitializeAddIn();
                    _isPaneInitialized = true;
                }
            }
        }
        return GOWordAgentPane;
    }
}
```

#### 3.2.6 新增 WpfFloatingWindow.cs（WPS 浮动窗口）

```csharp
// WpfFloatingWindow.cs — WPS 侧边栏替代方案
using System;
using System.Windows;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// WPS 环境下的浮动窗口，替代 CustomTaskPane
    /// 
    /// 特性：
    /// - 可拖动、可调大小
    /// - 始终置顶（跟随 WPS）
    /// - 内部复用 GOWordAgentPaneWpf 控件（零 UI 改动）
    /// - 记住上次位置和大小
    /// </summary>
    public class WpfFloatingWindow : Window
    {
        private GOWordAgentPaneWpf _wpfControl;

        public bool Visible
        {
            get => this.IsVisible;
            set
            {
                if (value) { this.Show(); this.Activate(); }
                else { this.Hide(); }
            }
        }

        public WpfFloatingWindow()
        {
            // 窗口属性
            this.Title = "智能校对";
            this.Width = 400;
            this.Height = 600;
            this.MinWidth = 300;
            this.MinHeight = 400;
            this.WindowStyle = WindowStyle.ToolWindow;
            this.ResizeMode = ResizeMode.CanResize;
            this.ShowInTaskbar = false;

            // 居中到屏幕右侧（模拟侧边栏位置）
            this.Left = SystemParameters.PrimaryScreenWidth - this.Width - 20;
            this.Top = 100;

            // 加载位置记录
            LoadPosition();

            // 内容：复用现有 WPF 控件！
            _wpfControl = new GOWordAgentPaneWpf();
            this.Content = _wpfControl;

            // 保存位置
            this.LocationChanged += SavePosition;
            this.SizeChanged += (s, e) => SavePosition();

            // 关闭时隐藏而非真正关闭
            this.Closing += (s, e) =>
            {
                e.Cancel = true;
                this.Hide();
            };
        }

        private void SavePosition()
        {
            try
            {
                string path = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "GOWordAgentAddIn", "wps-window-pos.txt");
                System.IO.File.WriteAllText(path,
                    $"{this.Left},{this.Top},{this.Width},{this.Height}");
            }
            catch { }
        }

        private void LoadPosition()
        {
            try
            {
                string path = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "GOWordAgentAddIn", "wps-window-pos.txt");
                if (System.IO.File.Exists(path))
                {
                    var parts = System.IO.File.ReadAllText(path).Split(',');
                    if (parts.Length == 4)
                    {
                        this.Left = double.Parse(parts[0]);
                        this.Top = double.Parse(parts[1]);
                        this.Width = double.Parse(parts[2]);
                        this.Height = double.Parse(parts[3]);
                    }
                }
            }
            catch { }
        }
    }
}
```

#### 3.2.7 Ribbon 处理

WPS 不支持 VSTO Ribbon 扩展，但 WPS 有自己的 Ribbon API。改动 `gowordagentribbon.cs`：

```csharp
// gowordagentribbon.cs — 改动

private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
{
    var addIn = ThisAddIn.Current;
    if (addIn == null) return;

    var pane = addIn.GetOrInitializePane();
    if (pane == null) return;

    if (pane is Microsoft.Office.Tools.CustomTaskPane ctp)
    {
        // Word: 切换 CustomTaskPane
        ctp.Visible = !ctp.Visible;
    }
    else if (pane is WpfFloatingWindow fw)
    {
        // WPS: 切换浮动窗口
        fw.Visible = !fw.Visible;
    }
}
```

**WPS 环境下 Ribbon 按钮的替代方案**：WPS 支持通过配置 XML 自定义 Ribbon，或通过 WPS 的 JS API 添加按钮。但这不影响核心功能运行——即使没有 Ribbon 按钮，浮动窗口也可以通过 COM Add-in 的其他方式触发。

#### 3.2.8 COM Add-in 注册（替代 VSTO 加载）

WPS 不走 VSTO 加载机制，需要在注册表注册标准 COM Add-in：

```registry
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Kingsoft\Office\Wps\AddinsWizard\Addins\{YOUR-ADDIN-CLSID}]
"FriendlyName"="GOWordAgent 智能校对"
"Description"="AI 智能校对插件"
"LoadBehavior"=dword:00000003
"CommandLineSafe"=dword:00000001
```

或通过代码注册（在安装程序中）：

```csharp
// 注册 COM Add-in 到 WPS
static void RegisterWpsAddin()
{
    string wpsRoot = @"Software\Kingsoft\Office\Wps\AddinsWizard\Addins";
    using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
        $@"{wpsRoot}\{CLSID}"))
    {
        key.SetValue("FriendlyName", "GOWordAgent 智能校对");
        key.SetValue("Description", "AI 智能校对插件");
        key.SetValue("LoadBehavior", 3); // 3 = 加载时启动
    }
}
```

---

## 四、不需要改动的文件（直接复用）

| 文件 | 原因 |
|------|------|
| `GOWordAgentPaneWpf.xaml.cs` | WPF 控件，被 `WpfFloatingWindow` 和 `GOWordAgentPaneHost` 共用 |
| `GOWordAgentPaneWpf.xaml` | WPF 界面，不变 |
| `ProofreadService.cs` | 不直接依赖 Word COM，通过 IWordInterop 间接使用 |
| `ProofreadCacheManager.cs` | Windows 路径，WPS 也跑在 Windows 上 |
| `ConfigManager.cs` | DPAPI 加密，WPS 也跑在 Windows 上 |
| `ILLMService.cs` | 纯接口 |
| `DeepSeekService.cs` | 纯 HTTP |
| `GLMService.cs` | 纯 HTTP |
| `OllamaService.cs` | 纯 HTTP |
| `LLMServiceFactory.cs` | 纯工厂 |
| `HttpClientFactory.cs` | 纯 HTTP |
| `LLMRequestLogger.cs` | 纯日志 |
| `ProofreadIssueParser.cs` | 纯字符串 |
| `DocumentSegmenter.cs` | 纯文本处理 |
| `ProofreadResultRenderer.cs` | WPF 渲染，不变 |
| `ViewModels/*.cs` | WPF MVVM，不变 |
| `Models/ProofreadModels.cs` | 纯数据模型 |

**共 17 个文件不需要改动，直接复用。**

---

## 五、必须实测验证的 WPS COM 差异

### 5.1 验证清单

| API | Word 行为 | WPS 预期 | 风险 |
|-----|----------|---------|------|
| `Find.Execute` Wrap 参数 | `wdFindStop` 枚举 | 数值 0 | 低（但需确认） |
| `TrackRevisions = true` 后 `range.Text = "new"` | 自动产生修订标记 | 可能直接替换不产生修订 | ⚠️ **高** |
| `Comments.Add(range, text)` | 第一个参数是 Range | 参数顺序可能不同 | ⚠️ 中 |
| `Range.Start` / `Range.End` | 返回 int | 可能返回 long | 低 |
| `ScrollIntoView(range)` | 滚动到位置 | 可能不支持 | 低（有降级） |
| `Selection.Text` | 获取选中 | 行为一致 | 低 |
| `Paragraphs.Count` | 段落数 | 行为一致 | 低 |

### 5.2 验证脚本（C#，在插件中运行）

```csharp
// WpsCompatTest.cs — 在 WPS 环境中运行的兼容性测试
// 可以做成插件内的隐藏菜单项触发

using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace GOWordAgentAddIn
{
    public static class WpsCompatTest
    {
        public static void RunAll()
        {
            dynamic app = null;
            dynamic doc = null;
            var sb = new System.Text.StringBuilder();

            try
            {
                app = Marshal.GetActiveObject("KWps.Application");
                doc = app.ActiveDocument;

                // Test 1: 获取文本
                sb.AppendLine($"[1] Content.Text 长度: {doc.Content.Text?.Length ?? -1}");

                // Test 2: Find
                dynamic range = doc.Content;
                dynamic find = range.Find;
                bool found = (bool)find.Execute("的", false, false, false, true, 0);
                sb.AppendLine($"[2] Find '的': {found}, pos={range.Start}-{range.End}");
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(range);

                // Test 3: TrackRevisions
                bool oldTrack = (bool)doc.TrackRevisions;
                doc.TrackRevisions = true;
                range = doc.Range(0, 3);
                string before = (string)range.Text;
                range.Text = "测试替换";
                string after = (string)range.Text;
                bool hasRevisions = doc.Revisions.Count > 0;
                // 撤销
                app.Undo(1);
                doc.TrackRevisions = oldTrack;
                sb.AppendLine($"[3] TrackRevisions: before='{before}', after='{after}', revisions={hasRevisions}");

                // Test 4: Comments
                range = doc.Range(0, 5);
                try
                {
                    var comment = doc.Comments.Add(range, "测试批注");
                    sb.AppendLine($"[4] Comments.Add: ✅ count={doc.Comments.Count}");
                    comment.Delete();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"[4] Comments.Add: ❌ {ex.Message}");
                }

                // Test 5: ScrollIntoView
                range = doc.Range(0, 10);
                range.Select();
                try
                {
                    app.ActiveWindow.ScrollIntoView(range);
                    sb.AppendLine("[5] ScrollIntoView: ✅");
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"[5] ScrollIntoView: ❌ {ex.Message}");
                }

                Marshal.ReleaseComObject(range);
            }
            catch (Exception ex)
            {
                sb.AppendLine($"[ERROR] {ex.Message}");
            }
            finally
            {
                if (doc != null) Marshal.ReleaseComObject(doc);
                if (app != null) Marshal.ReleaseComObject(app);
            }

            MessageBox.Show(sb.ToString(), "WPS COM 兼容性测试结果");
        }
    }
}
```

---

## 六、项目结构（改造后）

```
gowordagent/
├── Interop/                          # 【新增目录】
│   ├── IWordInterop.cs               # 【新增】COM 操作抽象接口
│   ├── HostType.cs                   # 【新增】宿主类型枚举
│   ├── HostDetector.cs               # 【新增】宿主检测
│   ├── InteropFactory.cs             # 【新增】统一工厂
│   ├── MsWordInterop.cs              # 【新增】原 WordDocumentService.cs 重构
│   └── WpsWordInterop.cs             # 【新增】WPS COM 适配
├── WpfFloatingWindow.cs              # 【新增】WPS 浮动窗口
├── WpsCompatTest.cs                  # 【新增】兼容性测试
│
├── ThisAddIn.cs                      # 【改动】增加 WPS 分支
├── WordProofreadController.cs        # 【改动】WordDocumentService → IWordInterop（~20行）
├── gowordagentribbon.cs              # 【改动】增加 WPS 窗口切换分支（~5行）
│
├── GOWordAgentPaneWpf.xaml.cs        # 【不动】
├── GOWordAgentPaneWpf.xaml           # 【不动】
├── GOWordAgentPaneHost.cs            # 【不动】
├── ProofreadService.cs               # 【不动】
├── ProofreadCacheManager.cs          # 【不动】
├── ConfigManager.cs                  # 【不动】
├── DocumentSegmenter.cs              # 【不动】
├── ProofreadIssueParser.cs           # 【不动】
├── ILLMService.cs                    # 【不动】
├── DeepSeekService.cs                # 【不动】
├── GLMService.cs                     # 【不动】
├── OllamaService.cs                  # 【不动】
├── LLMServiceFactory.cs              # 【不动】
├── HttpClientFactory.cs              # 【不动】
├── LLMRequestLogger.cs               # 【不动】
├── ProofreadResultRenderer.cs        # 【不动】
├── Models/ProofreadModels.cs         # 【不动】
├── ViewModels/*.cs                   # 【不动】
├── gowordagentribbon.designer.cs     # 【不动】
│
└── WordDocumentService.cs            # 【保留原文件不删，作为 MsWordInterop 的源参考】
```

---

## 七、改动量统计

| 类型 | 文件数 | 改动行数（估算） | 说明 |
|------|--------|----------------|------|
| **新增文件** | 7 | ~500 行 | IWordInterop + 2个实现 + Factory + Detector + 浮动窗口 + 测试 |
| **改动文件** | 3 | ~50 行 | ThisAddIn + WordProofreadController + Ribbon |
| **不动文件** | 17 | 0 行 | 直接复用 |
| **合计** | 27 | ~550 行新增/改动 | 原项目 ~3910 行 |

**改动占比：约 14%**（其中大部分是新增文件，改动现有代码仅 ~50 行）

---

## 八、工期估算

| 阶段 | 工作项 | 工期 | 前置依赖 |
|------|--------|------|---------|
| **S0** | WPS COM 兼容性验证（运行测试脚本） | **1-2 天** | 需要 Windows + WPS Pro |
| **S1** | 抽象接口 IWordInterop + MsWordInterop（从原文件重构） | **1 天** | S0 |
| **S2** | WpsWordInterop 实现 + InteropFactory | **2 天** | S0 + S1 |
| **S3** | WordProofreadController 改造（~20 行） | **半天** | S1 + S2 |
| **S4** | WpfFloatingWindow + ThisAddIn WPS 分支 | **1 天** | S2 |
| **S5** | Ribbon 适配 + COM Add-in 注册表 | **1 天** | S4 |
| **S6** | 端到端测试（Word + WPS 双环境） | **2 天** | S5 |
| **合计** | | **7-10 天** | |

---

## 九、风险与降级

| 风险 | 概率 | 影响 | 降级方案 |
|------|------|------|---------|
| `TrackRevisions` 在 WPS 不生效 | 中 | 高（核心功能） | 降级为直接替换+批注（原代码已有此逻辑） |
| `Comments.Add` 参数顺序不同 | 中 | 中（降级后无批注） | 尝试两种参数顺序，均失败则跳过批注 |
| `Find.Execute` 行为差异 | 低 | 中 | 用 `Content.Text.IndexOf()` 纯 JS 替代 |
| WPS 版本间 COM API 不一致 | 中 | 高 | 针对目标 WPS 版本逐个测试 |
| COM Add-in 在 WPS 加载失败 | 低 | 高 | 检查注册表路径，确认 CLSID 正确 |
| 浮动窗口体验不如侧边栏 | — | 低（体验降级） | 不可接受则退回 JS 插件方案 |

---

## 十、方案对比总结

| | 本方案（WPS COM） | JS 插件方案（信创版） |
|---|---|---|
| **改动量** | ~550 行 / 7-10 天 | ~2350 行 / 6-8 周 |
| **改动现有代码** | ~50 行 | ~400 行重写 |
| **新增代码** | ~500 行 | ~2000 行 |
| **复用率** | ~85% | ~30% |
| **运行环境** | Windows only | Windows + Linux |
| **信创支持** | ❌ 不支持 | ✅ 支持 |
| **风险** | WPS COM API 差异 | WPS JS API 差异 + 前后端分离 |
| **适合场景** | 快速让 WPS Windows 用户用上 | 长期信创布局 |

---

*本方案基于 gowordagent 源码实际分析，建议先完成 S0 POC（1-2天）验证 WPS COM 兼容性后再推进。*
