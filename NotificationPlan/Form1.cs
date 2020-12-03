using Microsoft.Office.Interop.Outlook;
using NotificationPlan.Data;
using NotificationPlan.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;

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
            if (!Other.IsExecutOutlook())
            {
                MessageBox.Show("Outlook не запущен. Для продолжения работы запустите Outlook и повторите попытку");
                return;
            }
            int month = Other.GetMonthToInt(cbMonth.SelectedItem.ToString());
            var date = new DateTime(DateTime.Now.Year, month, 1);
            var list = GetListItemCalendarsOfEmploy(date);
            Other.AddItemCalendar(list);
            setContext.Settings.MonthAddedInCalendar = date.Month;
            setContext.Settings.YearAddedInCalendar = date.Year;
            setContext.SaveSettings();
            timer2.Start();
        }
        /// <summary>
        /// Переделать для автоматизации
        /// </summary>
        private void CheckNextMonthFileWorkPlan()
        {
            IsWorkSync = true;
            if (DateTime.Now.Day >= 25)
            {
                
                var NextMonth = DateTime.Now.AddMonths(1);
                // Проверка на следующий год
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


        private List<ItemCalendar> GetListItemCalendarsOfEmploy(DateTime month)
        {
            if (FileWork.GetFileName(month.Month,
                month.Year) == null) return null;

            var workPlans = WorkPlanList.GetPlan(month.Month,
                month.Year);

            //var itemCalendars = Converter.Convert(
            //    workPlans.Where(t => t.NameEmploy.Contains(setContext.Settings.NameEmploy))
            //        .Where(t => t.EndTO >= DateTime.Now).ToList()
            var itemCalendars = Converter.Convert(
            workPlans.Where(t => t.NameEmploy
            .Contains(setContext.Settings.NameEmploy))
            .ToList());
            lAddedItem.Text = "Добавлено " + itemCalendars.Count.ToString();
            return itemCalendars;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbMonth.SelectedIndex = 0;
            OnRewrite();
            //if (!setContext.Settings.IsFirstOpen)
            //    timer1.Enabled = true;
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
            //timer1.Stop();
            setContext.Settings.NameEmploy = tbName.Text;
            setContext.Settings.DayRemineder = (int)nudReminder.Value;
            setContext.Settings.IsFirstOpen = false;
            setContext.SaveSettings();
            //timer1.Start();
            
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
            e.Cancel = false;
            return;
            //Для автоматиации
            /*if (!IsClose)
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
            }*/
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            //timer1.Enabled = false;
            //CheckNextMonthFileWorkPlan();
            OnRewrite();
            //timer1.Enabled = true;
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IsClose = true;
            this.Close();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            lAddedItem.Text = "";
            timer2.Stop();
        }

        
    }
}
