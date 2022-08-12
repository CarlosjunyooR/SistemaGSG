using System;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class Splash : Form
    {
        public Splash()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value < 100)
            {
                progressBar1.Value = progressBar1.Value + 20;
            }
            else
            {
                timer1.Enabled = false;
                this.Visible = false;
            }
        }
    }
}
