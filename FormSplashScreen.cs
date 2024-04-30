using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelQuery
{
    public partial class FormSplashScreen : Form
    {
        public FormSplashScreen()
        {
            InitializeComponent();
            // 設置定時器，控制Splash Screen的顯示時間
            timer1 = new System.Windows.Forms.Timer();
            timer1.Interval = 1000; // 3秒
            timer1.Tick += Timer_Tick;
            timer1.Start();
        }
        public FormSplashScreen(string message) : this()
        {
            label1.Text = message;
        }

        // 定時器事件處理程序，當時間到達時關閉Splash Screen
        private void Timer_Tick(object sender, EventArgs e)
        {
            timer1.Stop(); // 停止計時器
            this.Close(); // 關閉Splash Screen
        }

    }
}
