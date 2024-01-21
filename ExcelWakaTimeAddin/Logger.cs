using System;
using System.IO;
using System.Windows.Forms;
using WakaTime.Shared.ExtensionUtils;

namespace WakaTime.ExtensionUtils
{
	public class Logger : ILogger
	{
		private readonly bool _isDebugEnabled;
		private readonly StreamWriter _writer;

		private readonly string filename = $"{AppDataDirectory}\\excel-wakatime.log";

        public Logger(string configFilepath)
		{
			try
			{
                var configFile = new ConfigFile(configFilepath);

                _isDebugEnabled = configFile.GetSettingAsBoolean("debug");

                _writer = new StreamWriter(File.Open(filename, FileMode.Append, FileAccess.Write, FileShare.ReadWrite));
            }
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		private static string AppDataDirectory
		{
			get
			{
				var roamingFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				var appFolder = Path.Combine(roamingFolder, "WakaTime");

				// Create folder if it does not exist
				if (!Directory.Exists(appFolder))
					Directory.CreateDirectory(appFolder);

				return appFolder;
			}
		}

		public void Debug(string msg)
		{
			if (!_isDebugEnabled)
				return;

			Log(LogLevel.Debug, msg);
		}

		public void Info(string msg)
		{
			Log(LogLevel.Info, msg);
		}

		public void Warning(string msg)
		{
			Log(LogLevel.Warning, msg);
		}

		public void Error(string msg, Exception ex = null)
		{
			var exceptionMessage = $"{msg}: {ex}";

			Log(LogLevel.HandledException, exceptionMessage);
		}

		private void Log(LogLevel level, string msg)
		{
			var now = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt", new System.Globalization.CultureInfo("en-US"));

			try
			{
				_writer.WriteLine("[Wakatime {0} {1}] {2}", Enum.GetName(level.GetType(), level), now, msg);
				_writer.Flush();
			}
			catch (Exception ex)
			{
                MessageBox.Show(ex.ToString(), $"Error writing to \"{filename}\"", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                MessageBox.Show(ex.ToString(), $"{Enum.GetName(level.GetType(), level)} - {msg}", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
		}

		public void Close()
		{
			_writer?.Close();
		}
	}
}