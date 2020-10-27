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
        public Form1()
        {
            InitializeComponent();
            setContext = new SettingsContext();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WorkPlanList ww = new WorkPlanList();
            var workPlans = ww.GetPlan(10);
            var itemCalendars = new List<ItemCalendar>();
            label1.Text = workPlans.Count(empl => (empl.NameEmploy.Contains(Const.NameEmploy))).ToString();
            itemCalendars = Other.Convert(workPlans.Where(t => t.NameEmploy.Contains(Const.NameEmploy)).Where(t => t.StartTO >= DateTime.Now).ToList());
            foreach (var item in itemCalendars)
            {
                listBox1.Items.Add(item.StartDateTime + item.Body);
            }

            Other.AddItemCalendar(itemCalendars);

        }
    }
}
