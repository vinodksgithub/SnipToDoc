using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;

namespace ScreenCaptureUtility
{
    public static class SaveLocationManager
    {
        private const string KeyName = "SaveFolder";

        public static string GetSaveFolder()
        {
            string folder = ConfigurationManager.AppSettings[KeyName];

            if (string.IsNullOrEmpty(folder))
            {
                // Default to Desktop if not set
                folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            try
            {
                // ✅ Auto-create if missing
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }
            catch
            {
                // ✅ Fallback if path is invalid or cannot be created
                folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            return folder;
        }

        public static void SetSaveFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("Folder path cannot be empty.");

            //  Auto-create if missing
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (config.AppSettings.Settings[KeyName] == null)
                config.AppSettings.Settings.Add(KeyName, folderPath);
            else
                config.AppSettings.Settings[KeyName].Value = folderPath;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static void PromptAndSetFolder()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select folder to save captures";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    SetSaveFolder(fbd.SelectedPath);
                    MessageBox.Show($"Save location set to:\n{fbd.SelectedPath}", "Location Updated");
                }
            }
        }
    }
}