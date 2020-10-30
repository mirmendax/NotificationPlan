using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NotificationPlan.Models;

namespace NotificationPlan.Data
{
    public class SettingsContext
    {
        public Settings Settings { get; set; }

        public SettingsContext()
        {
            Settings = new Settings();
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
                SaveSettings();
            }
        }

        public void SetDefaultSettings()
        {
            Settings.NameEmploy = Environment.UserName.Split(new []{' '})[0];
            Settings.DayRemineder = Const.DayReminder;
            Settings.PathToWorkPlan = Const.PathToWorkPlan;
            Settings.StartFileWorkPlanWord = Const.StartFileWorkPlanWord;
            Settings.EndFileWorkPlanWord = Const.EndFileWorkPlanWord;
            Settings.MonthAddedInCalendar = 1;
            Settings.YearAddedInCalendar = 1;
            Settings.LastSync = Converter.ConvertToDateSync(DateTime.Now);
            Settings.IsFirstOpen = true;
            Settings.NameCalendar = Const.NameCalendar;
            SaveSettings();
        }

        public void SaveSettings()
        {
            var strJson = JsonConvert.SerializeObject(Settings, Formatting.Indented);
            File.WriteAllText(Const.NameFileSettings, strJson);
        }
    }
}