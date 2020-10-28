using Microsoft.Office.Interop.Outlook;
using NotificationPlan.Data;
using NotificationPlan.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NotificationPlan
{
    public partial class Form1 : Form
    {
        public SettingsContext setContext;
        /// <summary>
        /// Сворачивание окна
        /// </summary>
        public bool IsOpen = true;
        
        private bool IsClose = false;
        /// <summary>
        /// Выполняется синхронизация
        /// </summary>
        private bool IsWorkSync = false;

        private void SetPositionForm()
        {
            this.StartPosition = FormStartPosition.Manual;
            Point pt = Screen.PrimaryScreen.WorkingArea.Location;
            pt.Offset(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height);
            pt.Offset(-this.Width, -this.Height);
            this.Location = pt;
        }

        public Form1()
        {
            InitializeComponent();
            setContext = new SettingsContext();
            SetPositionForm();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void CheckNextMonthFileWorkPlan()
        {
            IsWorkSync = true;
            if (DateTime.Now.Day >= 25)
            {
                // Проверка на следующий год
                var NextMonth = DateTime.Now.AddMonths(1);
                if (NextMonth.Month == 1)
                {
                    if (setContext.Settings.YearAddedInCalendar < NextMonth.Year)
                    {
                        if (FileWork.GetFileName(NextMonth.Month, NextMonth.Year) != null)
                        {
                            var itemCalendars = GetListItemCalendarsOfEmploy(NextMonth);
                            if (itemCalendars != null)
                            {
                                setContext.Settings.YearAddedInCalendar = NextMonth.Year;
                                setContext.Settings.MonthAddedInCalendar = NextMonth.Month;
                                setContext.SaveSettings();
                            }
                        }
                        
                        //Other.AddItemCalendar(itemCalendars);
                    }
                }//Конец проверки на следующий год
                else
                {
                    if (setContext.Settings.MonthAddedInCalendar < NextMonth.Month)
                    {
                        if (FileWork.GetFileName(NextMonth.Month, NextMonth.Year) != null)
                        {
                            var itemCalendars = GetListItemCalendarsOfEmploy(NextMonth);
                            if (itemCalendars != null)
                            {
                                setContext.Settings.MonthAddedInCalendar = NextMonth.Month;
                                setContext.SaveSettings();
                            }
                            
                            //Other.AddItemCalendar(itemCalendars);
                        }                        
                    }
                }
            }

            setContext.Settings.LastSync = Converter.ConvertToDateSync(DateTime.Now);
            setContext.SaveSettings();
            IsWorkSync = false;
        }


        private List<ItemCalendar> GetListItemCalendarsOfEmploy(DateTime nextMonth)
        {
            if (FileWork.GetFileName(nextMonth.Month,
                nextMonth.Year) == null) return null;

            var workPlans = WorkPlanList.GetPlan(nextMonth.Month,
                nextMonth.Year);


            var itemCalendars = Converter.Convert(
                workPlans.Where(t => t.NameEmploy.Contains(setContext.Settings.NameEmploy))
                    .Where(t => t.EndTO >= DateTime.Now).ToList()
            );
            lAddedItem.Text = itemCalendars.Count.ToString();
            return itemCalendars;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OnRewrite();
            timer1.Enabled = true;
        }

        private void OnRewrite()
        {
            tbName.Text = setContext.Settings.NameEmploy;
            nudReminder.Value = setContext.Settings.DayRemineder;
            lLastSync.Text = "Последнее обновление: " + setContext.Settings.LastSync
                                                      + "\nЗагружен план работ на "
                                                      + Other.GetMonthToString(setContext.Settings
                                                          .MonthAddedInCalendar);
        }

        private void bSave_Click(object sender, EventArgs e)
        {
            setContext.Settings.NameEmploy = tbName.Text;
            setContext.Settings.DayRemineder = (int)nudReminder.Value;
            setContext.SaveSettings();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (IsOpen)
            {
                this.Hide();
            }
            else
            {
                this.Show();
                this.SetPositionForm();
                this.Activate();
            }

            IsOpen = !IsOpen;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!IsClose)
            {
                this.Hide();
                e.Cancel = true;
                return;
            }
            if (IsClose && !IsWorkSync)
            {
                e.Cancel = false;
                return;
            }
            else
            {
                e.Cancel = true;
                MessageBox.Show("Идет синхронизация. Дождитесь окончания, затем попробуйте снова.");
                return;
            }
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            //timer1.Enabled = false;
            CheckNextMonthFileWorkPlan();
            OnRewrite();
            //timer1.Enabled = true;
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IsClose = true;
            this.Close();
        }
    }
}
