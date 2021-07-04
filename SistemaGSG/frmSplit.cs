using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using iTextSharp.text.pdf;
using iTextSharp.text;
using java.io;
using System.IO;
using MetroFramework;

namespace SistemaGSG
{
    public partial class frmSplit : MetroFramework.Forms.MetroForm
    {
        public frmSplit()
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
            //PDDocument doc = PDDocument.load(txtUrl.Text);
            //PDFTextStripper stripper = new PDFTextStripper();
            //richTextBox1.Text = (stripper.getText(doc));
            ExtracaoPDFMetodo(new int[] { 3,7,22,53 }) ;
            MetroMessageBox.Show(this,"Finalizado com sucesso.","Aviso!");
        }
        static void ExtracaoPDFMetodo(int[] paginas)
        {
            PdfReader pdfReader = new PdfReader(@"C:\Users\junio\Desktop\1479067_725_06_2020_TODAS.PDF");
            Document document = new Document();

            if(pdfReader.NumberOfPages > 0)
            {
                foreach (var item in paginas)
                {
                    PdfCopy pdfCopy = new PdfCopy(document, new FileStream(Path.Combine(@"C:\Users\junio\Desktop\Extraidos\", string.Format("pagina_{0}.pdf",item)), FileMode.Create));
                    document.Open();
                    pdfCopy.AddPage(pdfCopy.GetImportedPage(pdfReader,item));
                }
                document.Close();
            }
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
                if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    frm_Main frm_Main = new frm_Main();
                    frm_Main.Show();
                    Close();
                }
        }
    }
}
