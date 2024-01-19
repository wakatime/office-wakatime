using System;
using System.Reflection;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using WakaTime.Forms;
using WakaTime.Shared.ExtensionUtils;
using WakaTime.ExtensionUtils;

namespace WordWakaTimeAddin
{
    public partial class WakaTimeAddIn
    {
        internal static SettingsForm SettingsForm;
        internal static WakaTime.Shared.ExtensionUtils.WakaTime WakaTime;

        private async void WakaTimeAddin_Startup(object sender, System.EventArgs e)
        {
            var metadata = new Metadata
            {
                EditorName = "word",
                PluginName = "word-wakatime",
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
                Application.WindowActivate += ApplicationOnWindowActivate;
                Application.DocumentOpen += ApplicationOnDocumentOpen;
                Application.DocumentBeforeSave += ApplicationOnDocumentBeforeSave;

                await WakaTime.InitializeAsync();
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("Error Initializing WakaTime", ex);
            }
        }

        #region Event Handlers

        private static void ApplicationOnWindowActivate(Word.Document doc, Word.Window wn)
        {
            try
            {
                WakaTime.HandleActivity(doc.FullName, false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnWindowActivate", ex);
            }
        }

        private static void ApplicationOnDocumentOpen(Word.Document doc)
        {
            try
            {
                WakaTime.HandleActivity(doc.FullName, false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnDocumentOpen", ex);
            }
        }

        private static void ApplicationOnDocumentBeforeSave(Word.Document doc, ref bool saveasui, ref bool cancel)
        {
            try
            {
                WakaTime.HandleActivity(doc.FullName, true, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnDocumentBeforeSave", ex);
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

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += WakaTimeAddin_Startup;
        }

        #endregion

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
}
