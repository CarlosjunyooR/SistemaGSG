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
    public partial class frm_Main : MetroFramework.Forms.MetroForm
    {
        string usuarioLogado = System.Environment.UserName;
        string nomeMaquina = System.Environment.MachineName;
        string dominio = System.Environment.UserDomainName;
        public frm_Main()
        {
            InitializeComponent();
            label9.Text = version;
            NomePC.Text = nomeMaquina;
            NomeDominio.Text = dominio;
            NomeUser.Text = usuarioLogado;
        }

        [assembly: AssemblyVersion("1.*")]
        string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        private void novaContaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ceal mont = new Ceal();
            mont.Show();
            this.Visible = false;
        }

        private void porCódÚnicoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormRel relMont = new FormRel();
            relMont.Show();
            this.Visible = false;
        }

        private void btnSair2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void testeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new FormNotificacao().Show();
            new FormNotific().Show();
        }

        private void frm_Main_Load(object sender, EventArgs e)
        {
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            dataHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            dataHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
        }

        private void criarPedidoSAPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormPedido pedidoSAP = new FormPedido();
            pedidoSAP.Show();
            this.Visible = false;
        }

        private void notasFiscaisFabianaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormNotaFiscal NotaFiscal = new FormNotaFiscal();
            NotaFiscal.Show();
            this.Visible = false;
        }

        private void pDFParaTXTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmPDF pdfTotxt = new frmPDF();
            pdfTotxt.Show();
            this.Visible = false;
        }

        private void pDFSepararToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmSplit ExtrairPDF = new frmSplit();
            ExtrairPDF.Show();
            this.Visible = false;
        }
    }
}
