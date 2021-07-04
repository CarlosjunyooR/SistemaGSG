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
using SistemaGSG;

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
