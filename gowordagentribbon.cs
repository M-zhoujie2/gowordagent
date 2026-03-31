using Microsoft.Office.Tools.Ribbon;

namespace GOWordAgentAddIn
{
    public partial class GOWordAgentRibbon
    {
        private void GOWordAgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = ThisAddIn.Current;
            if (addIn == null)
                return;

            // 使用按需初始化获取面板
            var pane = addIn.GetOrInitializePane();
            if (pane == null)
                return;

            pane.Visible = !pane.Visible;
        }
    }
}
