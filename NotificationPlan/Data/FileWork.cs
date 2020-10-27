using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NotificationPlan.Data
{
    public class FileWork
    {
        /// <summary>
        /// Получение имени файла из каталога с планами работ на месяц month
        /// </summary>
        /// <param name="month"></param>
        /// <returns></returns>
        public static string GetFileName(byte month, int year)
        {
            var path = Const.PathToWorkPlan+"\\"+year.ToString();
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
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.InitialDirectory = Const.PathToWorkPlan;
                    DialogResult dialogResult = ofd.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        fileMonth = ofd.FileName;
                    }
                }
            }
            

            return fileMonth;
        }


    }
}
