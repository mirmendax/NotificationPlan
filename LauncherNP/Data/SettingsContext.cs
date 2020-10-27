using LauncherNP.Models;
using NotificationPlan.Data;

namespace LauncherNP.Data
{
    public class SettingsContext
    {
        public Settings Settings { get; set; }

        public void SetDefaultSettings()
        {
            Settings.FileInfo = "";
            Settings.Version = "0";
            Settings.PathToUpdate = "";
        }
    }
}