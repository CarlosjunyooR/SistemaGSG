using MySql.Data.MySqlClient;
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
    public partial class FichaCadastroCliente : MetroFramework.Forms.MetroForm
    {
        public FichaCadastroCliente(string cCliente)
        {
            InitializeComponent();
            CCliente = cCliente;
            ConsultaFichaDeCadastroCliente();
        }

        private void ConsultaFichaDeCadastroCliente()
        {
            try
            {
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = ConexaoDados.GetConnectionFaturameto();
                MyCommand.CommandText = "SELECT * FROM `tb_cliente` WHERE col_cod_id="+CCliente+";";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    lblNomeCliente.Text = dreader["col_nome"].ToString();
                    lblRuaCliente.Text = dreader["col_rua"].ToString();
                    lblCidade.Text = dreader["col_cidade"].ToString();
                    lblCNPJ.Text = dreader["col_cnpj"].ToString();
                    break;
                }
                ConexaoDados.GetConnectionFaturameto().Close();
            }
            catch (MySqlException ErroDB)
            {
                MessageBox.Show("Banco de Dados, Desligado! " + ErroDB, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }

        public string CCliente { get; }
    }
}
