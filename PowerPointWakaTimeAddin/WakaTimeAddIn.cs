using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WakaTime.Forms;
using WakaTime.Shared.ExtensionUtils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using WakaTime.ExtensionUtils;

namespace PowerPointWakaTimeAddin
{
    public partial class WakaTimeAddIn
    {
        internal static SettingsForm SettingsForm;
        internal static WakaTime.Shared.ExtensionUtils.WakaTime WakaTime;

        private async void WakaTimeAddIn_Startup(object sender, EventArgs e)
        {
            var metadata = new Metadata
            {
                EditorName = "powerpoint",
                PluginName = "powerpoint-wakatime",
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
                Application.PresentationBeforeClose += ApplicationOnPresentationBeforeClose;                
                Application.PresentationOpen += ApplicationOnPresentationOpen;
                Application.PresentationSave += ApplicationOnPresentationSave;
                Application.WindowActivate += ApplicationOnWindowActivate;
                Application.SlideShowBegin += ApplicationOnSlideShowBegin;
                Application.SlideShowEnd += ApplicationOnSlideShowEnd;
                Application.SlideShowOnNext += ApplicationOnSlideShowOnNext;
                Application.SlideShowOnPrevious += ApplicationOnSlideShowOnPrevious;

                await WakaTime.InitializeAsync();
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
                WakaTime.HandleActivity(wn.Presentation.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(wn.Presentation.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(pres.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(wn.Presentation.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(pres.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(pres.FullName, true, string.Empty);
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
                WakaTime.HandleActivity(pres.FullName, false, string.Empty);
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
                WakaTime.HandleActivity(pres.FullName, false, string.Empty);
            }
            catch (Exception ex)
            {
                WakaTime.Logger.Error("ApplicationOnPresentationBeforeClose", ex);
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
