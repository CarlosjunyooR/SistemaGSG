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
        int attempt = 1;
        public frmLogin()
        {
            InitializeComponent();
            label3.Text = version;
        }
        [assembly: AssemblyVersion("1.*")]
        string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
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
                    label5.Visible = true;
                    label5.Text = "Erro você ainda tem " + attempt++ + " de 3";
                    label5.ForeColor = Color.Red;
                    //MessageBox.Show("Usuário ou Senha, Incorretos!");
                    FMP = false;
                    txtUser.Text = "";
                    txtSenha.Text = "";
                }
                if(attempt == 4)
                {
                    label6.Visible = true;
                    label6.Text = "Você teve " + attempt++ + " de 3 tentativas, Feche o programa e tente novamente.";
                    label6.ForeColor = Color.Blue;

                    txtUser.Visible = false;
                    label5.Visible = false;
                    txtSenha.Visible = false;
                    btnEntrar.Visible = false;
                    lblUsuario.Visible = false;
                    lblSenha.Visible = false;
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Olá Srº(a), " + txtUser.Text + " selecione uma conexão abaixo, para iniciar a\naplicação!.");
                gpBoxConexao.Focus();
            }
            catch (MySqlException)
            {
                MessageBox.Show("Olá Srº(a), " + txtUser.Text + " esta conexão encontra-se fechada para esta\naplicação! tente outra.");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (attempt<4)
            {
                logar();
            }
            else
            {

            }
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
                CONEX = new MySqlConnection(@"server=usga-servidor-m;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                txtConexao.Text = "usga-servidor-m";
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
                CONEX = new MySqlConnection(@"server=localhost;database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                txtConexao.Text = "localhost";
        }
        private void frmLogin_Load(object sender, EventArgs e)
        {
            //dateTimePicker1.Value = dateTimePicker1.Value.AddDays(30);
        }
    }
}
