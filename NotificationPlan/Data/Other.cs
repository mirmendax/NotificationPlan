using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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
            /*switch (month)
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
                default: result = string.Empty;
                    break;

            }*/
            if (month < 1 || month > 12) return string.Empty;
            result = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
            return result;
        }
        public static int GetMonthToInt(string strMonth)
        {
            if (string.IsNullOrEmpty(strMonth)) return 0;
            var result = 0;
            var cultureInfo = new CultureInfo("ru-RU");
            var date = $"1 {strMonth} 2020";
            var dateFormat = DateTime.Parse(date, cultureInfo);
            result = dateFormat.Month;
            /*switch (strMonth)
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
            }*/
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

        public static bool IsExecuteOutlook()
        {
            return Process.GetProcessesByName("Outlook").Any();

        }

        #region Not_used
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
        

        #endregion
        
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
        /// Получение имени файла из каталога с планами работ на месяц month
        /// </summary>
        /// <param name="month"></param>
        /// <param name="year"></param>
        /// <returns></returns>
        public static string GetFileName(int month, int year)
        {
            var setContext = new SettingsContext();
            
            var path = setContext.Settings.PathToWorkPlan+"\\"+year.ToString();
            var strMonth = Other.GetMonthToString(month);
            string fileMonth = null;
            if (Directory.Exists(path))
            {
                var files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    if (file.Contains(strMonth))
                    {
                        fileMonth = file;
                        break;
                    }
                }
                if (string.IsNullOrEmpty(fileMonth))
                {
                    //OpenFileDialog ofd = new OpenFileDialog();
                    //ofd.InitialDirectory = Const.PathToWorkPlan;
                    //DialogResult dialogResult = ofd.ShowDialog();
                    //if (dialogResult == DialogResult.OK)
                    //{
                    //    fileMonth = ofd.FileName;
                    //}
                }
            }
            
            return fileMonth;
        }

    }
}
