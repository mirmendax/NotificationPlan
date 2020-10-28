using System.IO;
using LauncherNP.Models;
using Newtonsoft.Json;


namespace LauncherNP.Data
{
    public class SetLauncherContext
    {
        public Settings Settings { get; set; }

        public SetLauncherContext()
        {
            LoadSettings();
        }
        
        private void LoadSettings()
        {
            Settings = new Settings();
            if (File.Exists(Const.SettingsLauncherFileName))
            {
                var fJson = File.ReadAllText(Const.SettingsLauncherFileName);
                Settings = JsonConvert.DeserializeObject<Settings>(fJson);
                
            }
            else
            {
                SetDefaultSettings();
                SaveSettings();
            }
        }

        private void SetDefaultSettings()
        {
            Settings.FileInfo = Const.Default.FileInfo;
            Settings.Build = Const.Default.Build;
            Settings.PathToUpdate = Const.Default.PathToUpdate;
            Settings.ExecFileName = Const.Default.ExecFile;
        }

        public void SaveSettings()
        {
            var strJson = JsonConvert.SerializeObject(Settings, Formatting.Indented);
            File.WriteAllText(Const.SettingsLauncherFileName, strJson);
        }
    }
}