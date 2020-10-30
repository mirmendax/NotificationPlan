using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using NotificationPlan.Models;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;

namespace NotificationPlan.Data
{
    public class Other
    {
        /// <summary>
        /// Получение строкового представления месяца (надо переделать)
        /// </summary>
        /// <param name="month"></param>
        /// <returns></returns>
        /// <exception cref="Exception">Если число 1<month>12 </exception>
        public static string GetMonthToString(int month)
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
        public static int GetMonthToInt(string strMonth)
        {
            var result = 0;
            switch (strMonth)
            {
                case "Январь": result = 1;
                    break;
                case "Февраль":
                    result = 2;
                    break;
                case "Март":
                    result = 3;
                    break;
                case "Апрель":
                    result = 4;
                    break;
                case "Май":
                    result = 5;
                    break;
                case "Июнь":
                    result = 6;
                    break;
                case "Июль":
                    result = 7;
                    break;
                case "Август":
                    result = 8;
                    break;
                case "Сентябрь":
                    result = 9;
                    break;
                case "Октябрь":
                    result = 10;
                    break;
                case "Ноябрь":
                    result = 11;
                    break;
                case "Декабрь":
                    result = 12;
                    break;

                default:
                    
                    break;
            }
            return result;
        }
        /// <summary>
        /// Добавление Событий в календарь
        /// </summary>
        /// <param name="listItem">список событий</param>
        public static void AddItemCalendar(List<ItemCalendar> listItem)
        {
            SettingsContext sContext = new SettingsContext();
            Application Application = null;
            Application = new Application();
            MAPIFolder primaryCalendar = Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            // if (!IsExistsWorkCalendar())
            // {
            //     CreateNewCalendar(sContext.Settings.NameCalendar);
            // }

            var personalCalendar = primaryCalendar.Folders[sContext.Settings.NameCalendar];
            if (personalCalendar == null) return;
            foreach (var item in listItem)
            {
                AppointmentItem newEvent = primaryCalendar.Items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
                if (newEvent != null)
                {
                    newEvent.Start = item.StartDateTime;
                    newEvent.End = item.EndDateTime;
                    //newEvent.AllDayEvent = true;
                    if (item.IsReminded)
                    {
                        newEvent.ReminderMinutesBeforeStart = item.ReminderMinute;
                        newEvent.ReminderPlaySound = true;
                    }
                    else
                    {
                        newEvent.ReminderSet = false;
                    }
                    
                    newEvent.Subject = item.Title;
                    newEvent.Body = item.Body;
                    newEvent.Save();
                }
            }
            Application.ActiveExplorer().CurrentFolder.Display();
        }

        public static bool IsExecutOutlook()
        {
            return Process.GetProcessesByName("Outlook").Any();

        }

        /*public static bool IsExistsWorkCalendar()
        {
            SettingsContext sContext = new SettingsContext();
            var result = false;
            Application App = new Application();
            var explorer = App.Application.ActiveExplorer();
            MAPIFolder primary = explorer.Session
                .GetDefaultFolder((OlDefaultFolders.olFolderCalendar));
            foreach (MAPIFolder folder in primary.Folders)
            {
                if (folder.Name == sContext.Settings.NameCalendar)
                {
                    result = true;
                    break;
                    
                }
            }

            return result;
        }
        */

        /*public static void CreateNewCalendar(string name)
        {
            Application App = new Application();
            var explorer = App.Application.ActiveExplorer();
            MAPIFolder primary = explorer.Session
                .GetDefaultFolder((OlDefaultFolders.olFolderCalendar));
            var needFolder = true;
            
            if (!IsExistsWorkCalendar())
            {
                primary.Folders.Add(name, OlDefaultFolders.olFolderCalendar);
                App.Application.ActiveExplorer().SelectFolder(primary.Folders[name]);
                App.Application.ActiveExplorer().CurrentFolder.Display();
            }
        }
        */
        /// <summary>
        /// Получение кол-во рабочих дней для получения уведомления
        /// </summary>
        /// <param name="startDate">Дата события</param>
        /// <returns></returns>
        public static int DayOfWeek(DateTime startDate, int dayReminder)
        {
            switch (startDate.AddDays(-dayReminder).DayOfWeek)
            {
                case System.DayOfWeek.Saturday:
                    return dayReminder + 1;
                case System.DayOfWeek.Sunday:
                    return dayReminder + 2;
                default:
                    return dayReminder;
            }
        }
        /// <summary>
        /// Конвертирование из плана в элемент события календаря с группировкой по дате
        /// </summary>
        /// <param name="wPlan"></param>
        /// <returns></returns>

    }
}
