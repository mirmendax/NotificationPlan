using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using NotificationPlan.Models;
using Exception = System.Exception;

namespace NotificationPlan.Data
{
    public class Other
    {
        public static void Test()
        {
            
        }
        public static string GetMonthToString(byte month)
        {
            string result = "";
            switch (month)
            {
                case 1: result = "Январь"; break;
                case 2: result = "Февраль"; break;
                case 3: result = "Март"; break;
                case 4: result = "Апрель"; break;
                case 5: result = "Май"; break;
                case 6: result = "Июнь"; break;
                case 7: result = "Июль"; break;
                case 8: result = "Август"; break;
                case 9: result = "Сентябрь"; break;
                case 10: result = "Октябрь"; break;
                case 11: result = "Ноябрь"; break;
                case 12: result = "Декабрь"; break;
                default: throw new Exception("Нет такого месяца");

            }
            return result;
        }

        public static void AddItemCalendar(List<ItemCalendar> listItem)
        {
            Application Application = null;
            Application = new Application();
            MAPIFolder primaryCalendar = Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            foreach (var item in listItem)
            {
                AppointmentItem newEvent = primaryCalendar.Items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
                if (newEvent != null)
                {
                    newEvent.Start = item.StartDateTime;
                    newEvent.End = item.EndDateTime;
                    //newEvent.AllDayEvent = true;
                    newEvent.ReminderMinutesBeforeStart = item.ReminderMinute;
                    newEvent.ReminderPlaySound = true;
                    newEvent.Subject = item.Title;
                    newEvent.Body = item.Body;
                    newEvent.Save();
                }
            }
            Application.ActiveExplorer().CurrentFolder.Display();
        }

        public static int DayOfWeek(DateTime startDate)
        {
            if (startDate.AddDays(-1).DayOfWeek == System.DayOfWeek.Saturday || 
                startDate.AddDays(-1).DayOfWeek== System.DayOfWeek.Sunday)
            {
                return 3;
            }
            else
            {
                return 1;
            }
        }
        public static List<ItemCalendar> Convert(List<WorkPlan> wPlan)
        {
            var result = new List<ItemCalendar>();

            var groupList = wPlan.GroupBy(startDate => startDate.StartTO);
            foreach (var itemGroup in groupList)
            {
                var temp = new ItemCalendar();
                foreach (var item in itemGroup)
                {
                    temp.Title = "Работа";
                    temp.Body += item.Title + " " +item.ViewTO + ";\n ";
                    temp.StartDateTime = item.StartTO.AddHours(8).AddMinutes(2).AddSeconds(17);
                    temp.EndDateTime = item.EndTO.AddHours(8).AddMinutes(3).AddSeconds(17);
                    temp.ReminderDay = DayOfWeek(item.StartTO);
                }
                result.Add(temp);
            }

            return result;
        }
    }
}
