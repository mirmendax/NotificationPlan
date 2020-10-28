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
using LauncherNP.Data;

namespace LauncherNP
{
    public partial class Form1 : Form
    {
        SetLauncherContext sLContext = new SetLauncherContext();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Updater.IsEnableUpdate().Item1)
            {
                label1.Text = "Найдено обновление";
                Updater.Update(Environment.CurrentDirectory + "\\");
                label1.Text = "Обновление завершено.";
            }
            
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var prcApp = new Process();
            prcApp.StartInfo.FileName = sLContext.Settings.ExecFileName;
            prcApp.Start();
            Close();
        }
    }
}