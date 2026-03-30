using System.Windows;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// 隐私同意对话框
    /// </summary>
    public partial class PrivacyConsentDialog : Window
    {
        /// <summary>
        /// 用户是否同意
        /// </summary>
        public bool IsAgreed { get; private set; } = false;

        public PrivacyConsentDialog()
        {
            InitializeComponent();
        }

        private void BtnAgree_Click(object sender, RoutedEventArgs e)
        {
            IsAgreed = true;
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            IsAgreed = false;
            DialogResult = false;
            Close();
        }
    }
}
