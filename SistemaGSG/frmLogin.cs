using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Odbc;
using MySql.Data.MySqlClient;
using MySql.Data;

namespace SistemaGSG
{
    public partial class frmLogin : MetroFramework.Forms.MetroForm
    {
        public frmLogin()
        {
            InitializeComponent();
        }

        public bool FMP = false;

        public void logar()
        {
            try
            {
                string tb_user = "SELECT * FROM tb_gsg WHERE nome = @usuario";
                MySqlCommand cmd;
                MySqlDataReader dr;

                cmd = new MySqlCommand(tb_user, CONEX);

                //Verificar Usuário//
                cmd.Parameters.Add(new MySqlParameter("@usuario", txtUser.Text));
                CONEX.Open();
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                while (dr.Read())
                {
                    dados.usuario = Convert.ToString(dr["nome"]);
                    dados.senha = Convert.ToString(dr["senha"]);
                    dados.nivel = Convert.ToInt32(dr["status"]);
                }
                CONEX.Close();

                if (dados.senha == txtSenha.Text)
                {
                    FMP = true;
                    this.Dispose();

                }
                else
                {
                    MessageBox.Show("Usuário ou Senha, Incorretos!");
                    FMP = false;

                    txtUser.Text = "";
                    txtSenha.Text = "";
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Está Maquina Não Tem Acesso a Esta Conexão\nTente Outra!");
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            logar();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        MySqlConnection CONEX;
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
                CONEX = new MySqlConnection(@"server=10.2.1.95;database=ceal1;Uid=remoto;Pwd=MbunHhYiRffEMAtl;");
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
                CONEX = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
        }
        private void frmLogin_Load(object sender, EventArgs e)
        {

        }
    }
}
