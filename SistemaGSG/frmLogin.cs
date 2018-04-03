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

                 SqlConnection conn;
                 SqlCommand cmd;
                 SqlDataReader dr;

                //Conexão com o Banco de Dados//
                conn = new SqlConnection(conexao.conex);

                //Conectando a Tabela dentro do Banco//
                cmd = new SqlCommand(tb_user, conn);

                //Verificar Usuário//
                cmd.Parameters.Add(new SqlParameter("@usuario", txtUser.Text));
                conn.Open();
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                while (dr.Read())
                {
                    dados.usuario = Convert.ToString(dr["nome"]);
                    dados.senha = Convert.ToString(dr["senha"]);
                    dados.nivel = Convert.ToInt32(dr["nivel"]);
                }
                conn.Close();

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
            
            catch(Exception err)
            {
                MessageBox.Show("Erro!" + err);
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

        private void frmLogin_Load(object sender, EventArgs e)
        {

        }
    }
}
