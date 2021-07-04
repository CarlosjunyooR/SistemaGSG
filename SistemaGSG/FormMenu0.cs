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
    public partial class FormMenu0 : MetroFramework.Forms.MetroForm
    {
        public FormMenu0()
        {
            InitializeComponent();
            txtCaminho.Visible = false;
            btnImportar.Visible = false;
            lblCaminho.Visible = false;
        }

        private void importarArquivoCNAB240ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtCaminho.Visible == true)
            {
                txtCaminho.Visible = false;
                btnImportar.Visible = false;
                lblCaminho.Visible = false;
            }
            else
            {
                txtCaminho.Visible = true;
                btnImportar.Visible = true;
                lblCaminho.Visible = true;
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtCaminho.Text))
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Arquivos RET(*.ret)|*.ret";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string path = ofd.FileName.ToString();
                    txtCaminho.Text = path;
                }
            }
        }
    }
}
