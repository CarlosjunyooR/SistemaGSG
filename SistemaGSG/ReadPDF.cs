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
    public partial class ReadPDF : MetroFramework.Forms.MetroForm
    {
        public ReadPDF()
        {
            InitializeComponent();
        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog ofd=new OpenFileDialog() { ValidateNames = true, Multiselect=false, Filter = "PDF|*.pdf" })
            {
                if (ofd.ShowDialog()==DialogResult.OK)
                {
                    axAcroPDF1.src = ofd.FileName;
                }
            }
        }
    }
}
