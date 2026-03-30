using Microsoft.Office.Tools.Ribbon;
using System;

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
            if (addIn == null || addIn.GOWordAgentPane == null)
                return;

            addIn.GOWordAgentPane.Visible = !addIn.GOWordAgentPane.Visible;
        }
    }
}
