using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PowerPointWakaTimeAddin.Forms
{
    public partial class ApiKeyForm : Form
    {
        private readonly WakaTime.Shared.ExtensionUtils.WakaTime _wakaTime;

        public ApiKeyForm(ref WakaTime.Shared.ExtensionUtils.WakaTime wakaTime)
        {
            _wakaTime = wakaTime;
            InitializeComponent();
        }

        private void ApiKeyForm_Load(object sender, EventArgs e)
        {
            try
            {
                txtAPIKey.Text = _wakaTime.Config.ApiKey;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }            
        }
        
        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                var matched = Regex.IsMatch(txtAPIKey.Text.Trim(), "(?im)^(waka_)?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}$");                              
                if (matched)
                {
                    _wakaTime.Config.ApiKey = txtAPIKey.Text.Trim();
                    _wakaTime.Config.Save();
                    _wakaTime.Config.ApiKey = txtAPIKey.Text.Trim();
                }
                else
                {
                    MessageBox.Show(@"Please enter valid Api Key.");
                    DialogResult = DialogResult.None; // do not close dialog box
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
