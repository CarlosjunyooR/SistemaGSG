using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormCadastroFornecedor : MetroFramework.Forms.MetroForm
    {
        public FormCadastroFornecedor(string text)
        {
            InitializeComponent();
            txtCodigoFornec.Text = text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_fornecedor (id_CodFornecedor, col_Nome) " +
                    "VALUE" +
                    "S ('" + txtCodigoFornec.Text + "', '" + txtNomeFornec.Text + "')", ConexaoDados.GetConnectionFornecedor());
                prompt_cmd.ExecuteNonQuery();
                ConexaoDados.GetConnectionFornecedor().Close();

                MessageBox.Show("Cadastrado!");
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }
    }
}
