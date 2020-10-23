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
        private List<WorkPlan> _planList = new List<WorkPlan>();

        public List<WorkPlan> GetPlan(byte month)
        {
            var strFile = FileWork.GetFileName(month);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<WorkPlan> workPlans = new List<WorkPlan>();
            int endWorkCount = 0;
            int startWorkCount = 0;
            if (!string.IsNullOrEmpty(strFile))
            {
                using (var excel = new ExcelPackage(new FileInfo(strFile)))
                {
                    var sheet = excel.Workbook.Worksheets["ТАиВ"];
                    if (sheet != null)
                    {
                        for (int i = 1; i < sheet.Cells.Count(); i++)
                        {
                            var value = (sheet.Cells[$"B{i}"].Value as string);
                            if (value == null) continue;
                            
                            if (value.Contains(Const.StartFileWorkPlanWord)) 
                                startWorkCount = i + 1;
                            
                            if (value.Contains(Const.EndFileWorkPlanWord) || value == "")
                            {
                                endWorkCount = i;
                                break;
                            }
                        }

                        for (int i = startWorkCount; i < endWorkCount; i++)
                        {
                            var workPlan = new WorkPlan();
                            workPlan.Title = sheet.Cells[i, 2].Value as string;
                            workPlan.ViewTO = sheet.Cells[i, 3].Value as string;
                            workPlan.NameEmploy = sheet.Cells[i, 11].Value as string;
                            
                            try
                            {
                                workPlan.StartTO = DateTime.Parse(sheet.Cells[i, 4].Value.ToString());
                                workPlan.EndTO = DateTime.Parse(sheet.Cells[i, 5].Value.ToString());
                            }
                            catch (Exception e)
                            {
                                continue;
                            }
                            workPlans.Add(workPlan);
                        }
                    }
                }
            }

            return workPlans;
        }
    }
}
