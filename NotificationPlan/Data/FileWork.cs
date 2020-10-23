using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NotificationPlan.Data
{
    public class FileWork
    {
        public static string GetFileName(byte month)
        {
            var path = Const.PathToWorkPlan;
            var strMonth = Other.GetMonthToString(month);
            string fileMonth = null;
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
                //
            }

            return fileMonth;
        }
    }
}
