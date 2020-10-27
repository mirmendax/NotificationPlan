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
        public bool IsOpen = true;
        private bool IsClose = false;

        private void SetPositionForm()
        {
            this.StartPosition = FormStartPosition.Manual;
            //Верхний левый угол экрана
            Point pt = Screen.PrimaryScreen.WorkingArea.Location;
            //Перенос в нижний правый угол экрана без панели задач
            pt.Offset(Screen.PrimaryScreen.WorkingArea.Width, Screen.PrimaryScreen.WorkingArea.Height);
            //Перенос в местоположение верхнего левого угла формы, чтобы её правый нижний угол попал в правый нижний угол экрана
            pt.Offset(-this.Width, -this.Height);
            //Новое положение формы
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
            WorkPlanList ww = new WorkPlanList();
            var workPlans = ww.GetPlan(10);
            var itemCalendars = new List<ItemCalendar>();
            
            itemCalendars = Other.Convert(
                workPlans.Where(t => t.NameEmploy.Contains(setContext.Settings.NameEmploy))
                .Where(t => t.EndTO >= DateTime.Now).ToList()
                );
            label1.Text = itemCalendars.Count.ToString();

            Other.AddItemCalendar(itemCalendars);

        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
            tbName.Text = setContext.Settings.NameEmploy;
            nudReminder.Value = setContext.Settings.DayRemineder;
            

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
            }
            else
            {
                e.Cancel = false;
                
            }
        }



        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IsClose = true;
            this.Close();
        }
    }
}
