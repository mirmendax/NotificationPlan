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
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application Application = null;
            Application = new Microsoft.Office.Interop.Outlook.Application();
            MAPIFolder primaryCalendar = (MAPIFolder)Application.ActiveExplorer().Session
                .GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
           //var personalCalendar = primaryCalendar.Folders.Add("newCalendarName", OlDefaultFolders.olFolderCalendar);
            AppointmentItem newEvent = primaryCalendar.Items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
            newEvent.Start = DateTime.Now.AddHours(1).AddDays(1);
            newEvent.End = DateTime.Now.AddHours(1.25).AddDays(1);
            newEvent.Subject = "New Plan";
            newEvent.Body = "Mesf sdf dskpfldskfnlsd ";
            newEvent.ReminderMinutesBeforeStart = 1140;
            newEvent.ReminderPlaySound = true;
            newEvent.Save();
            Application.ActiveExplorer().CurrentFolder.Display();
        }
    }
}
