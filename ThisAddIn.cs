using System;
using System.IO;

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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Current = this;
            
            _paneHost = new GOWordAgentPaneHost();
            GOWordAgentPane = CustomTaskPanes.Add(_paneHost, "智能校验");
            GOWordAgentPane.Visible = true;
            
            // 加载保存的宽度
            var savedWidth = LoadSavedPaneWidth() ?? DefaultPaneWidth;
            GOWordAgentPane.Width = savedWidth;
            _cachedPaneWidth = savedWidth;

            // 宽度变更时实时保存（使用 _paneHost 的 SizeChanged 事件）
            _paneHost.SizeChanged += (s, args) =>
            {
                _cachedPaneWidth = GOWordAgentPane.Width;
                SavePaneWidthSafe(_cachedPaneWidth);
            };
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Shutdown 时不再访问 CustomTaskPane，使用缓存值
            SavePaneWidthSafe(_cachedPaneWidth);
        }

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
