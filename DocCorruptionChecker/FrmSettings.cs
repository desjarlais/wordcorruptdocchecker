using System.Windows.Forms;

namespace DocCorruptionChecker
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();

            ckRemoveFallback.Checked = Properties.Settings.Default.RemoveFallback == "true";
            ckOpenInWord.Checked = Properties.Settings.Default.OpenInWord == "true";
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.RemoveFallback = ckRemoveFallback.Checked ? "true" : "false";
            Properties.Settings.Default.OpenInWord = ckOpenInWord.Checked ? "true" : "false";
            Close();
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}