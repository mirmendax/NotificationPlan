using NotificationPlan.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using OfficeOpenXml;

namespace NotificationPlan.Data
{
    public class WorkPlanList
    {
        

        public static List<WorkPlan> GetPlan(int month, int year)
        {
            var strFile = FileWork.GetFileName(month, year);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var workPlans = new List<WorkPlan>();
            if (string.IsNullOrEmpty(strFile)) return workPlans;
            using (var excel = new ExcelPackage(new FileInfo(strFile)))
            {
                var sheet = excel.Workbook.Worksheets["ТАиВ"];
                if (sheet == null) return workPlans;
                var endWorkCount = 0;
                var startWorkCount = 0;
                (startWorkCount, endWorkCount) = StartWorkCount(sheet);
                var endDayMonth = new DateTime(year, month + 1, 1).AddDays(-1).Day;

                for (var i = startWorkCount; i < endWorkCount; i++)
                {
                    var workPlan = new WorkPlan();
                    workPlan.Title = (sheet.Cells[i, 2].Value as string) == null
                        ? ""
                        : (sheet.Cells[i, 2].Value as string);
                    
                    workPlan.ViewTO = (sheet.Cells[i, 3].Value as string) == null
                        ? ""
                        : (sheet.Cells[i, 3].Value as string);

                    workPlan.NameEmploy = (sheet.Cells[i, 11].Value as string) == null
                        ? ""
                        : (sheet.Cells[i, 11].Value as string);
                            
                    try
                    {
                        var dStart = (sheet.Cells[i, 4].Value.ToString()) == null
                            ? new DateTime(year, month, 1).ToString("d")
                            : sheet.Cells[i, 4].Value.ToString();
                        
                        
                        workPlan.StartTO = DateTime.Parse(dStart);
                        
                        var dEnd = (sheet.Cells[i, 5].Value.ToString()) == null
                            ? new DateTime(year, month, endDayMonth).ToString("d")
                            : sheet.Cells[i, 5].Value.ToString();
                        
                        workPlan.EndTO = DateTime.Parse(dEnd);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    workPlans.Add(workPlan);
                }
            }

            return workPlans;
        }

        private static (int startWorkCount, int endWorkCount) StartWorkCount(ExcelWorksheet sheet)
        {
            int startWorkCount = 0;
            int endWorkCount = 0;
            var setContext = new SettingsContext();
            for (int i = 1; i < sheet.Cells.Count(); i++)
            {
                var value = (sheet.Cells[i, 2].Value as string);
                if (value == null) continue;

                if (value.Contains(setContext.Settings.StartFileWorkPlanWord))
                    startWorkCount = i + 1;

                if (value.Contains(setContext.Settings.EndFileWorkPlanWord) || value == "" ||
                    value.Contains(setContext.Settings.EndFileWorkPlanWord.ToUpper()))
                {
                    endWorkCount = i;
                    break;
                }
            }

            return (startWorkCount, endWorkCount);
        }
    }
}
