using System;
using System.Collections.Generic;
using System.Linq;
using NotificationPlan.Models;

namespace NotificationPlan.Data
{
    public class Converter
    {
        public static List<ItemCalendar> ConvertAndGroup(List<WorkPlan> wPlan)
        {
            var result = new List<ItemCalendar>();

            var groupList = wPlan.GroupBy(startDate => startDate.StartTO);
            foreach (var itemGroup in groupList)
            {
                var temp = new ItemCalendar();
                foreach (var item in itemGroup)
                {
                    temp.Title = "Работа";
                    temp.Body += item.Title + " " + item.ViewTO + ";\n";
                    temp.StartDateTime = item.StartTO.AddHours(8).AddMinutes(2).AddSeconds(17);
                    temp.EndDateTime = item.EndTO.AddHours(8).AddMinutes(3).AddSeconds(17);
                    temp.ReminderDay = Other.DayOfWeek(item.StartTO, Const.DayReminder);
                }

                result.Add(temp);
            }

            return result;
        }

        public static DateSync ConvertToDateSync(DateTime date)
        {
            var result = new DateSync();
            result.Day = date.Day;
            result.Month = date.Month;
            result.Year = date.Year;
            result.Hour = date.Hour;
            result.Minute = date.Minute;
            return result;
        }

        public static DateTime ConvertToDateTime(DateSync date)
        {
            return new DateTime(date.Year, date.Month, date.Day, date.Hour, date.Minute, 0);
        }

        public static List<ItemCalendar> Convert(List<WorkPlan> wPlan)
        {
            var result = new List<ItemCalendar>();
            foreach (var item in wPlan)
            {
                var temp = new ItemCalendar();
                temp.Title = item.Title;
                temp.IsReminded = !string.IsNullOrEmpty(item.ViewTO);
                temp.Body = item.Title + " " + item.ViewTO;
                temp.StartDateTime = item.StartTO.AddHours(8);
                temp.EndDateTime = item.EndTO.AddHours(15);
                temp.ReminderDay = Other.DayOfWeek(item.StartTO, new SettingsContext().Settings.DayRemineder);
                result.Add(temp);
            }

            return result;
        }
    }
}