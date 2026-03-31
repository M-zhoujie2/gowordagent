using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media;

namespace GOWordAgentAddIn
{
    public partial class ThisAddIn
    {
        internal static ThisAddIn Current;
        internal Microsoft.Office.Tools.CustomTaskPane GOWordAgentPane;
        private GOWordAgentPaneHost _paneHost;

        private const string SettingsDir = "SmartProofreadingAddIn";
        private const string SettingsFile = "paneWidth.txt";
        private const int DefaultPaneWidth = 400;

        // 缓存当前面板宽度
        private int _cachedPaneWidth = DefaultPaneWidth;
        
        // 标记面板是否已初始化
        private bool _isPaneInitialized = false;
        private readonly object _initLock = new object();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Current = this;
            
            // 极简启动：不立即初始化面板，只在用户第一次使用时才创建
            // 这样可以最大程度避免拖慢 Word 启动速度
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] 插件已加载，面板将在首次使用时初始化");
        }
        
        /// <summary>
        /// 获取或初始化任务面板（按需延迟初始化）
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane GetOrInitializePane()
        {
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
        
        /// <summary>
        /// 延迟初始化插件（避免阻塞 Word 启动）
        /// </summary>
        private void InitializeAddIn()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] 开始初始化面板...");
                
                // 配置 WPF DPI 感知
                ConfigureDpiAwareness();
                
                _paneHost = new GOWordAgentPaneHost();
                GOWordAgentPane = CustomTaskPanes.Add(_paneHost, "智能校验");
                
                // 初始状态设为不可见，用户可以通过 Ribbon 按钮手动打开
                GOWordAgentPane.Visible = false;
                
                // 加载保存的宽度
                var savedWidth = LoadSavedPaneWidth() ?? DefaultPaneWidth;
                GOWordAgentPane.Width = savedWidth;
                _cachedPaneWidth = savedWidth;

                // 宽度变更时实时保存
                _paneHost.SizeChanged += (s, args) =>
                {
                    if (GOWordAgentPane != null)
                    {
                        _cachedPaneWidth = GOWordAgentPane.Width;
                        SavePaneWidthSafe(_cachedPaneWidth);
                    }
                };
                
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] 面板初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThisAddIn] 初始化失败: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Shutdown 时不再访问 CustomTaskPane，使用缓存值
            SavePaneWidthSafe(_cachedPaneWidth);
            
            // 释放 PaneHost
            if (_paneHost is IDisposable disposable)
            {
                disposable.Dispose();
            }
        }

        #region DPI 感知配置

        /// <summary>
        /// 配置 WPF DPI 感知 - 修复 MaterialDesign 在高分辨率屏幕的模糊问题
        /// </summary>
        private static void ConfigureDpiAwareness()
        {
            try
            {
                // 设置进程级 DPI 感知（如果尚未设置）
                // 注意：对于 VSTO 外接程序，DPI 感知主要由宿主应用程序（Word）决定
                // 但我们可以通过 WPF 的 API 优化渲染
                
                // 禁用 WPF 的 DPI 缩放（让 Windows 处理）
                // 这可以防止 MaterialDesign 控件在高 DPI 下出现模糊
                // 注意：RenderMode 不是公开 API，跳过此设置
                System.Diagnostics.Debug.WriteLine($"[ThisAddIn] 渲染层级: {RenderCapability.Tier >> 16}");
                
                // 设置 TextOptions 以改善文本清晰度
                TextOptions.TextFormattingModeProperty.OverrideMetadata(
                    typeof(Window),
                    new FrameworkPropertyMetadata(TextFormattingMode.Display));
                
                TextOptions.TextHintingModeProperty.OverrideMetadata(
                    typeof(Window),
                    new FrameworkPropertyMetadata(TextHintingMode.Auto));
                
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] DPI 感知配置完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThisAddIn] DPI 感知配置失败: {ex.Message}");
                // 不影响主功能
            }
        }

        #endregion

        #region VSTO 生成的代码
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        #region 设置持久化

        private static string GetSettingsPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), SettingsDir);
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, SettingsFile);
        }

        private static int? LoadSavedPaneWidth()
        {
            try
            {
                string path = GetSettingsPath();
                if (File.Exists(path) && int.TryParse(File.ReadAllText(path), out int width) && width > 0)
                    return width;
            }
            catch { /* 忽略读取错误 */ }
            return null;
        }

        private static void SavePaneWidthSafe(int width)
        {
            if (width <= 0) return;
            try { File.WriteAllText(GetSettingsPath(), width.ToString()); }
            catch { /* 忽略写入错误 */ }
        }

        #endregion
    }
}
