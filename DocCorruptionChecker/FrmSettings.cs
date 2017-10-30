using System.Windows.Forms;

namespace DocCorruptionChecker
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();

            if (Properties.Settings.Default.RemoveFallback.ToString() == "true")
            {
                ckRemoveFallback.Checked = true;
            }
            else
            {
                ckRemoveFallback.Checked = false;
            }

            if (Properties.Settings.Default.OpenInWord == "true")
            {
                ckOpenInWord.Checked = true;
            }
            else
            {
                ckOpenInWord.Checked = false;
            }
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            if (ckRemoveFallback.Checked)
            {
                Properties.Settings.Default.RemoveFallback = "true";
            }
            else
            {
                Properties.Settings.Default.RemoveFallback = "false";
            }

            if (ckOpenInWord.Checked)
            {
                Properties.Settings.Default.OpenInWord = "true";
            }
            else
            {
                Properties.Settings.Default.OpenInWord = "false";
            }

            Close();
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}