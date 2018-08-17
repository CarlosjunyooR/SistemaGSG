using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class frmProtocolo : Form
    {
        public frmProtocolo()
        {
            InitializeComponent();
        }

        private void txtUrl_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                webBrowser.Navigate(txtUrl.Text);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            webBrowser.GoBack();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            webBrowser.GoForward();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            webBrowser.Refresh();
        }

        private void btnIr_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUrl.Text))
                webBrowser.Navigate(txtUrl.Text);
        }
    }
}
