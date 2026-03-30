using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace GOWordAgentAddIn
{
    public partial class GOWordAgentPaneHost : UserControl
    {
        private ElementHost _host;
        public GOWordAgentPaneWpf WpfControl { get; private set; }

        public GOWordAgentPaneHost()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = WpfControl = new GOWordAgentPaneWpf()
            };
            Controls.Add(_host);
        }
    }
}
