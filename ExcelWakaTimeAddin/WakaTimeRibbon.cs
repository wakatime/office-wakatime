using System;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelWakaTimeAddin
{
    public partial class WakaTimeRibbon
    {
        private void WakaTimeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                WakaTimeAddIn.SettingsForm.ShowDialog();
            }
            catch (Exception ex)
            {
                WakaTimeAddIn.WakaTime.Logger.Error("btnSettings_Click", ex);
            }
        }
    }
}
