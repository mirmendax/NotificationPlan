using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NotificationPlan.Models;

namespace NotificationPlan.Data
{
    public class SettingsContext
    {
        public Settings Settings = new Settings();

        public SettingsContext()
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            if (File.Exists(Const.NameFileSettings))
            {
                var fJson = File.ReadAllText(Const.NameFileSettings);
                Settings = JsonConvert.DeserializeObject<Settings>(fJson);
                
            }
            else
            {
                SetDefaultSettings();
            }
        }

        private void SetDefaultSettings()
        {
            Settings.NameEmploy = "Прокопенко";
            Settings.DayRemineder = Const.DayReminder;
            Settings.PathToWorkPlan = Const.PathToWorkPlan;
            Settings.StartFileWorkPlanWord = Const.StartFileWorkPlanWord;
            Settings.EndFileWorkPlanWord = Const.EndFileWorkPlanWord;
            Settings.MonthAddedInCalendar = DateTime.Now.Month;
            SaveSettings();
        }

        public void SaveSettings()
        {
            var strJson = JsonConvert.SerializeObject(Settings, Formatting.Indented);
            File.WriteAllText(Const.NameFileSettings, strJson);
        }
    }
}