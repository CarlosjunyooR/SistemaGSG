using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class ConfigConexao : MetroFramework.Forms.MetroForm
    {
        public ConfigConexao()
        {
            InitializeComponent();
        }
        //VariaveisDeclaradas
        string conexaoBancoDados;
        private void btnLocalizar_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Arquivos de Dados (*.accdb)|*.ACCDB";
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                txtCaminho.Text = ofd.FileName;
            }
        }

        private void btnVisualizarSenha_Click(object sender, EventArgs e)
        {
            if(txtSenhaDB.PasswordChar == '*')
            {
                txtSenhaDB.PasswordChar = default;
            }
            else
            {
                txtSenhaDB.PasswordChar = '*';
            }
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            try
            {
                conexaoBancoDados = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Jet OLEDB:Database Password={1};",txtCaminho.Text, txtSenhaDB.Text);
                if (string.IsNullOrEmpty(txtCaminho.Text))
                {

                }
                else
                {
                    AppSetting setting = new AppSetting();
                    setting.SaveConnectionString("DBSGSG", conexaoBancoDados);
                    if(MessageBox.Show("Sua conexão foi salva com sucesso!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        //Application.Restart();
                        //Environment.Exit(0);


                    }
                }
            }
            catch (OleDbException eRR)
            {
                MessageBox.Show(eRR.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ConfigConexao_Load(object sender, EventArgs e)
        {

        }

        private void btnFechar_Click(object sender, EventArgs e)
        {
            //this.Close();
            if (ConexaoBancoDeDadosOffline.DBSGSG_Conex().State == System.Data.ConnectionState.Open)
            {
                MessageBox.Show("Conectado!");
            }
            else
            {
                MessageBox.Show("Não Conectado!, Verificar...");
            }
        }
    }
}
