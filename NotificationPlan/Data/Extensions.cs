using System;
using System.Collections;
using System.Collections.Generic;
using NotificationPlan.Models;

namespace NotificationPlan.Data
{
    public static class Extensions
    {
        public static List<ItemCalendar> ConvertToItemCalendar<T>(this T self)
        where T: List<WorkPlan>, new()
        {
            var result = new List<ItemCalendar>();
            foreach (var item in self)
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

        /*public static List<ItemCalendar> ConvertAndGroup<T>(this T self)
        where T: List<WorkPlan>, new()
        {
            
        */
        //}
    }
}