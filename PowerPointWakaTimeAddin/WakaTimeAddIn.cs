using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using PowerPointWakaTimeAddin.Forms;
using WakaTime.Forms;
using WakaTime.Shared.ExtensionUtils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointWakaTimeAddin
{
    public partial class WakaTimeAddIn
    {
        internal static SettingsForm SettingsForm;
        internal static WakaTime.Shared.ExtensionUtils.WakaTime WakaTime;

        private void WakaTimeAddIn_Startup(object sender, EventArgs e)
        {
            var configuration = new Configuration
            {
                EditorName = "powerpoint",
                PluginName = "powerpoint-wakatime",
                EditorVersion = Application.Version,
                PluginVersion = Constants.PluginVersion
            };

            WakaTime = new WakaTime.Shared.ExtensionUtils.WakaTime(null, configuration, new Logger());

            WakaTime.Logger.Debug("Initializing in background thread.");
            Task.Run(() => { InitializeAsync(); }).ContinueWith(t => OnStartupComplete());
        }

        private void InitializeAsync()
        {
            try
            {
                // Settings Form
                SettingsForm = new SettingsForm(ref WakaTime);
                SettingsForm.ConfigSaved += SettingsFormOnConfigSaved;

                // setup event handlers
                Application.PresentationBeforeClose += ApplicationOnPresentationBeforeClose;
                Application.PresentationOpen += ApplicationOnPresentationOpen;
                Application.PresentationSave += ApplicationOnPresentationSave;
                Application.WindowActivate += ApplicationOnWindowActivate;
                Application.SlideShowBegin += ApplicationOnSlideShowBegin;
                Application.SlideShowEnd += ApplicationOnSlideShowEnd;
                Application.SlideShowOnNext += ApplicationOnSlideShowOnNext;
                Application.SlideShowOnPrevious += ApplicationOnSlideShowOnPrevious;

                WakaTime.InitializeAsync();
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("Error Initializing WakaTime", ex);
            }
        }

        #region Event Handlers

        private static void ApplicationOnSlideShowOnPrevious(PowerPoint.SlideShowWindow wn)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wn.Presentation.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnSlideShowOnPrevious", ex);
            }
        }

        private static void ApplicationOnSlideShowOnNext(PowerPoint.SlideShowWindow wn)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wn.Presentation.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnSlideShowOnNext", ex);
            }
        }

        private static void ApplicationOnSlideShowEnd(PowerPoint.Presentation pres)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(pres.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnSlideShowEnd", ex);
            }
        }

        private static void ApplicationOnSlideShowBegin(PowerPoint.SlideShowWindow wn)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(wn.Presentation.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnSlideShowBegin", ex);
            }
        }

        private static void ApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(pres.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWindowActivate", ex);
            }
        }

        private static void ApplicationOnPresentationSave(PowerPoint.Presentation pres)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(pres.FullName), true, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnPresentationSave", ex);
            }
        }

        private static void ApplicationOnPresentationOpen(PowerPoint.Presentation pres)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(pres.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnPresentationOpen", ex);
            }
        }

        private static void ApplicationOnPresentationBeforeClose(PowerPoint.Presentation pres, ref bool cancel)
        {
            try
            {
                WakaTime.HandleActivity(GetFileLocalPath(pres.FullName), false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnPresentationBeforeClose", ex);
            }
        }

        private static void OnStartupComplete()
        {
            // Prompt for api key if not already set
            if (string.IsNullOrEmpty(WakaTime.Config.ApiKey))
                PromptApiKey();
        }

        private static void SettingsFormOnConfigSaved(object sender, EventArgs eventArgs)
        {
            WakaTime.Config.Read();
        }

        #endregion

        #region Methods        

        private static void PromptApiKey()
        {
            WakaTime.Logger.Info("Please input your api key into the wakatime window.");

            var form = new ApiKeyForm(ref WakaTime);
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
