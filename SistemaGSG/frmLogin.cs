using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using MetroFramework;
using MySql.Data.MySqlClient;


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
        string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
        public bool FMP = false;
        
        public void logar()
        {
            try
            {
                string tb_user = "SELECT * FROM tb_user WHERE nome = @usuario";
                MySqlCommand cmd;
                MySqlDataReader dr;
                cmd = new MySqlCommand(tb_user, ConexaoDados.GetConnectionEquatorial());

                //Verificar Usuário//
                cmd.Parameters.Add(new MySqlParameter("@usuario", txtUser.Text));
                
                dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                while (dr.Read())
                {
                    dados.usuario = Convert.ToString(dr["nome"]);
                    dados.senha = Convert.ToString(dr["senha"]);
                    dados.nivel = Convert.ToInt32(dr["nivel"]);
                }
                ConexaoDados.GetConnectionEquatorial().Close();
                if (dados.senha == txtSenha.Text)
                {
                    if (dados.nivel == 3)
                    {
                        FormRelacao AbrirForm = new FormRelacao();
                        AbrirForm.Show();
                    }
                    if(dados.nivel==1)
                    {
                        FMP = true;
                        this.Dispose();
                    }
                }
                else
                {
                    int cont = 3;
                    int Menos = cont - attempt;
                    attempt++;
                    label5.Visible = true;
                    label5.Text = "Erro você ainda tem " + Menos + " chances.";
                    label5.ForeColor = Color.Red;
                    FMP = false;
                    txtUser.Text = "";
                    txtSenha.Text = "";
                }
                if(attempt == 4)
                {
                    label6.Visible = true;
                    label6.Text = "Você teve 3 de 3 tentativas, Feche o programa e tente novamente.";
                    label6.ForeColor = Color.Blue;

                    txtUser.Visible = false;
                    label5.Visible = false;
                    txtSenha.Visible = false;
                    btnEntrar.Visible = false;
                    lblUsuario.Visible = false;
                    lblSenha.Visible = false;
                }
            }
            catch (MySqlException ErrO)
            {
                MessageBox.Show("Erro no Banco de Dados! - \n Não Foi Possivel Conectar!");
                if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Application.Exit();
                }
                else
                {
                    if (MessageBox.Show("Deseja entrar no modo offline?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        FMP = true;
                        this.Dispose();
                    }
                }
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
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
                Application.Exit();
        }
        private void frmLogin_Load(object sender, EventArgs e)
        {
            try
            {
                if (ConexaoDados.GetConnectionEquatorial().State == ConnectionState.Open)
                {
                    label1.ForeColor = Color.Lime;
                    label1.Text = "Conectado...";
                }
                else
                {
                    label1.ForeColor = Color.Red;
                    label1.Text = "Não Conectado...";
                }
                ConexaoDados.GetConnectionEquatorial().Close();
            }
            catch(MySqlException MysqlErr)
            {
                MessageBox.Show("Erro no Banco de Dados! -\nNão Foi Possivel Conectar!");
                if (MessageBox.Show("Deseja encerrar a aplicação ?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Application.Exit();
                }
                label1.Text = "Não Conectado...";
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }
    }
}
