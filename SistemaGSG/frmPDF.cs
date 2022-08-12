using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class frmPDF : MetroFramework.Forms.MetroForm
    {
        public frmPDF()
        {
            InitializeComponent();
        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "PDF Files(*.pdf)|*.pdf";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string path = ofd.FileName.ToString();
                txtUrl.Text = path;
            }
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = "*.txt";
            saveFileDialog1.Filter = "TXT File|*.txt";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0)
            {
                richTextBox1.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.PlainText);
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            PDDocument doc = PDDocument.load(txtUrl.Text);
            PDFTextStripper stripper = new PDFTextStripper();
            richTextBox1.Text = (stripper.getText(doc));
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main frm_Main = new frm_Main();
                frm_Main.Show();
                this.Visible = false;
            }
        }
    }
}
