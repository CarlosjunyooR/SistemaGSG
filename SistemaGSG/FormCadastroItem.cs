using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace SistemaGSG
{
    public partial class FormCadastroItem : MetroFramework.Forms.MetroForm
    {
        public FormCadastroItem()
        {
            InitializeComponent();
        }
        public FormCadastroItem(string conexao)
        {
            InitializeComponent();
            txtHost.Text = conexao;
        }

        private void FormCadastroItem_Load(object sender, EventArgs e)
        {

        }
        MySqlConnection CONEXAOBD;
        MySqlCommand cmd;
        private void btnGravar_Click(object sender, EventArgs e)
        {
            try
            {
                CONEXAOBD = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                CONEXAOBD.Open();

                cmd = new MySqlCommand("INSERT INTO `tb_produtos` (`CD_PRODUTO`, `DESC_PRODUTO`, `CD_SAP`) VALUES ('" + txtCodProduto.Text + "', '" + txtDescProduto.Text + "', '" + txtCodSap.Text + "')", CONEXAOBD);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Salvo com Sucesso!", "Informação",MessageBoxButtons.OK,MessageBoxIcon.Information);

                CONEXAOBD.Close();

                txtCodProduto.Text = "";
                txtDescProduto.Text = "";
                txtCodSap.Text = "";

            }
            catch (MySqlException Err)
            {
                MessageBox.Show(Err.Message);
            }catch(Exception Erro)
            {
                MessageBox.Show(Erro.Message);
            }
        }
    }
}
