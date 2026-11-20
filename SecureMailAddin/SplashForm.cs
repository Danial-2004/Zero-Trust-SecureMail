using System;
using System.Windows.Forms;

namespace SecureMailAddin
{
    public partial class SplashForm : Form
    {
        public SplashForm()
        {
            InitializeComponent();
            this.Load += SplashForm_Load;
            timer1.Tick += Timer1_Tick;
        }

        private void SplashForm_Load(object sender, EventArgs e)
        {
            // Start the timer to close splash automatically
            timer1.Start();
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            this.Close(); // Auto close splash after interval
        }
    }
}
