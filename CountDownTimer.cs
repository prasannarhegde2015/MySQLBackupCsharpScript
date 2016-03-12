using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mystopwathc
{
    public partial class Form1 : Form
    {
       
        string strhh, strmm, strss;
        
        int ms = 0; int ss = 0; int mm = 0; int hh = 0;
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            textBox1.ReadOnly = true;
            textBox1.Text = "00:00:00";
           
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
           
                timer1.Start();
                timer1.Tick+=timer1_Tick;
                button1.Enabled = false;
        }

        public string gettimenow()
        {
                ms = ms + 100;
                System.Threading.Thread.Sleep(100);
                if (ms == 1000)
                {
                    ss = ss + 1;
                    ms = 0;
                }
                if (ss == 60)
                {
                    mm = mm + 1;
                    ss = 0;
                }
                if (mm == 60)
                {
                    hh = hh + 1;
                    mm = 0;
                }
                if (hh == 24)
                {
                    hh = 0;
                }
                if (hh < 10)
                {
                    strhh = "0" + hh.ToString();
                }
                else
                {
                    strhh = hh.ToString();
                }
                if (mm < 10)
                {
                    strmm = "0" + mm.ToString();
                }
                else
                {
                    strmm = mm.ToString();
                }
                if (ss < 10)
                {
                    strss = "0" + ss.ToString();
                }
                else
                {
                    strss = ss.ToString();
                }

                return strhh + ":" + strmm + ":" + strss;
           
        }
        public void timer1_Tick(object sender, EventArgs e)
        {
            textBox1.Text = gettimenow();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            button1.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            textBox1.Text = "00:00:00";
            button1.Enabled = true;
             ms = 0;  ss = 0;  mm = 0;  hh = 0;
        }
    }
}
