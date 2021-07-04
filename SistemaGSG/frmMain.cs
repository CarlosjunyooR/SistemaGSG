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
        string nomeMaquina = System.Environment.MachineName;
        string dominio = System.Environment.UserDomainName;
        public frm_Main()
        {
            InitializeComponent();
            label9.Text = version;
            NomePC.Text = nomeMaquina;
            NomeDominio.Text = dominio;
            NomeUser.Text = dados.usuario;
            dados.tema = Style.ToString();
        }

        [assembly: AssemblyVersion("1.*")]
        string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        private void novaContaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1)
            {
                Ceal AbrirForm = new Ceal();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!","Informação",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void porCódÚnicoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1)
            {
                FormRel AbrirForm = new FormRel();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        private void timer1_Tick(object sender, EventArgs e)
        {
            dataHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            dataHora.Text = (DateTime.Now.ToString("dd/MM/yy HH:mm:ss"));
        }
        private void notasFiscaisFabianaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel==1)
            {
                FormNotaFiscal AbrirForm = new FormNotaFiscal();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pDFParaTXTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel==3 || dados.nivel==4)
            {
                frmPDF AbrirForm = new frmPDF();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pDFSepararToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1)
            {
                frmSplit AbrirForm = new frmSplit();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void importarXMLSAPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                frmXML AbrirForm = new frmXML();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void posiçãoDaSemanaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                frmPosicaoSemana AbrirForm = new frmPosicaoSemana();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void baixaEmRequisiçõesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void fornecedorDeCanaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1)
            {
                FormDesconto AbrirForm = new FormDesconto();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void importarXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                frmXML AbrirForm = new frmXML();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void consultarNotasEmitidasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                frmProtocolo AbrirForm = new frmProtocolo();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void relaçãoDeNotasFiscaisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                FormRelacao AbrirForm = new FormRelacao();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void baixarXMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                FormDownloadXML AbrirForm = new FormDownloadXML();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void criarPedidoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                FormPedido AbrirForm = new FormPedido();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void criarAcessoBalançaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dados.nivel == 1 || dados.nivel == 3 || dados.nivel == 4)
            {
                FormAcesso AbrirForm = new FormAcesso();
                AbrirForm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Sem Autorização!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
