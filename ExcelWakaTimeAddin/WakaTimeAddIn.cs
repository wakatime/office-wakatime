using System;
using System.Reflection;
using System.Threading.Tasks;
using WakaTime.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WakaTime.Shared.ExtensionUtils;
using WakaTime.ExtensionUtils;

namespace ExcelWakaTimeAddin
{
    public partial class WakaTimeAddIn
    {
        internal static SettingsForm SettingsForm;
        internal static WakaTime.Shared.ExtensionUtils.WakaTime WakaTime;

        private async void WakaTimeAddIn_Startup(object sender, EventArgs e)
        {
            var metadata = new Metadata
            {
                EditorName = "excel",
                PluginName = "excel-wakatime",
                EditorVersion = Application.Version,
                PluginVersion = Constants.PluginVersion
            };

            WakaTime = new WakaTime.Shared.ExtensionUtils.WakaTime(metadata, new Logger(Dependencies.GetConfigFilePath()));

            WakaTime.Logger.Debug("Initializing in background thread.");

            await InitializeAsync();

            // Prompt for api key if not already set
            if (string.IsNullOrEmpty(WakaTime.Config.GetSetting("api_key")))
                PromptApiKey();
        }

        private async Task InitializeAsync()
        {
            try
            {
                // Settings Form
                SettingsForm = new SettingsForm(WakaTime.Config, WakaTime.Logger);

                // setup event handlers                
                Application.WorkbookOpen += ApplicationOnWorkbookOpen;
                Application.WorkbookAfterSave += ApplicationOnWorkbookAfterSave;
                Application.WindowActivate += ApplicationOnWindowActivate;
                Application.WorkbookActivate += ApplicationOnWorkbookActivate;

                await WakaTime.InitializeAsync();
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("Error Initializing WakaTime", ex);
            }
        }

        #region Event Handlers

        private static void ApplicationOnWorkbookActivate(Excel.Workbook wb)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wb.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWorkbookActivate", ex);
            }
        }

        private static void ApplicationOnWindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wb.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWindowActivate", ex);
            }
        }

        private static void ApplicationOnWorkbookAfterSave(Excel.Workbook wb, bool success)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wb.FullName), true, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWorkbookAfterSave", ex);
            }
        }

        private static void ApplicationOnWorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wb.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWorkbookOpen", ex);
            }
        }

        #endregion

        #region Methods        

        private static void PromptApiKey()
        {
            WakaTime.Logger.Info("Please input your api key into the wakatime window.");

            var form = new ApiKeyForm(WakaTime.Config, WakaTime.Logger);
            form.ShowDialog();
        }

        private static string GetFileLocalPath(string rawPath)
        {
            try
            {
                var workbookUri = new Uri(rawPath);
                // If local file, return it as-is
                if (workbookUri.IsFile)
                    return rawPath;

                string localPath = string.Empty;
                // Registry key names to loop                
                System.Collections.Generic.List<string> keyNames = new System.Collections.Generic.List<string>() { "OneDrive", "OneDriveCommercial", "OneDriveConsumer", "onedrive" };
                foreach (var keyName in keyNames)
                {
                    using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Environment", false))
                    {
                        var rootDirectory = key.GetValue(keyName).ToString();
                        if (string.IsNullOrEmpty(rootDirectory) == false && System.IO.Directory.Exists(rootDirectory))
                        {
                            if (keyName == "OneDrive")
                            {
                                int index = workbookUri.LocalPath.IndexOf('/', 1);
                                if (index > 0)
                                {
                                    string subPath = workbookUri.LocalPath.Substring(index);
                                    localPath = string.Format("{0}{1}", rootDirectory, subPath);
                                    localPath = localPath.Replace("/", "\\"); // slashes in the right direction
                                    if (System.IO.File.Exists(localPath))
                                        return localPath;
                                }
                            }

                            var pathParts = new System.Collections.Generic.Queue<string>(workbookUri.LocalPath.Split('/'));
                            do
                            {
                                // Compose a local path by adding root directory and slashes in between                                
                                string subPath = string.Join("\\", pathParts);
                                if (subPath[0] != '\\')
                                {
                                    subPath = "\\" + subPath;
                                }
                                localPath = string.Format("{0}{1}", rootDirectory, subPath);

                                if (System.IO.File.Exists(localPath))
                                    return localPath;

                                // The file was not found - get rid of leftmost part of the path and try again
                                pathParts.Dequeue();
                            }
                            while (pathParts?.Count > 0);

                        }
                    }
                }
                return localPath;
            }
            catch (Exception ex)
            {
                return rawPath;
            }
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += WakaTimeAddIn_Startup;
        }

        #endregion
    }

    public static class CoreAssembly
    {
        private static readonly Assembly Reference = typeof(CoreAssembly).Assembly;
        public static readonly Version Version = Reference.GetName().Version;
    }

    internal static class Constants
    {
        internal static readonly string PluginVersion =
            $"{CoreAssembly.Version.Major}.{CoreAssembly.Version.Minor}.{CoreAssembly.Version.Build}";
    }
}
