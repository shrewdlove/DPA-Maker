using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DPA_Maker
{

    public partial class loadingscreen1 : Form
    {
       int shouldIquit = 1;
       public void InitTimer()
        {
            timer1 = new Timer();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 2000; // in miliseconds
            timer1.Start();
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
                label2.Show();
                shouldIquit = shouldIquit + 1;
                if (shouldIquit == 3)
                {
                    this.Close();
                }
            
       
        }
        public loadingscreen1()
        {
            InitializeComponent();
            

            InitTimer();


        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        public Timer timer1 { get; set; }

        private void label1_Click_2(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
