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


namespace SistemaGSG
{
    public partial class FormNotificacao : MetroFramework.Forms.MetroForm
    {
        public FormNotificacao()
        {
            InitializeComponent();
        }

        private void FormNotificacao_Resize(object sender, EventArgs e)
        {
            //verifica se o formulario esta minimizado
            if (this.WindowState == FormWindowState.Minimized)
            {
                //esconde o formulário
                this.Hide();
                //deixa o aviso visivel
                notifyIcon1.Visible = true;
            }
        }

        private void contextMenuStrip1_Opening(object sender, EventArgs e)
        {
            //para abrir o formulário form1 mesmo, no seu caso, para você seria a tela de login ou o principal, ou outra tela mesmo
            new frm_Main().Show();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //irá exibir o formulário (neste caso o form1) em seu caso, pode ser a tela de login ou o principal, não sei como está a aplicação ai...
            this.Show();
            //o formulario irá iniciar maximizado
            this.WindowState = FormWindowState.Maximized;
            //oculta o aviso
            notifyIcon1.Visible = false;
        }

        //Classes de Datas
        Int32 segundos, minutos, milisegundos;
        DateTime dataHora = DateTime.Now;
        DateTime dataHora2 = DateTime.Now.AddDays(10);

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection con = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
                con.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", con);
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+10).ToShortDateString();

                dataHora = DataServidor;
                minutos = dataHora.Minute;
                segundos = dataHora.Second;
                milisegundos = dataHora.Millisecond;

                MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM contas WHERE data BETWEEN @DataServidor AND @dataFuturo", con);


                command.Parameters.AddWithValue("@dataFuturo", dataHora2);
                command.Parameters.AddWithValue("@DataServidor", dataHora);
                command.ExecuteNonQuery();

                int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
                con.Close();

                if (qtdVencer > 0)
                {
                    MessageBox.Show("Tem " + qtdVencer + " boletos pra vencer!");
                    FormNotificacao_Load(e, e);
                }
                else
                {
                    MessageBox.Show("Não tem boletos para vencer!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void FormNotificacao_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection(@"server=localhost;database=ceal1;Uid=root;Pwd=vertrigo;");
            con.Open();
            MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", con);
            DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
            string novadata = DataServidor.AddDays(+10).ToShortDateString();

            dataHora = DataServidor;
            minutos = dataHora.Minute;
            segundos = dataHora.Second;
            milisegundos = dataHora.Millisecond;

            MySqlCommand command = new MySqlCommand("SELECT COUNT(*) FROM contas WHERE data BETWEEN @DataServidor AND @dataFuturo", con);

            command.Parameters.AddWithValue("@dataFuturo", dataHora2);
            command.Parameters.AddWithValue("@DataServidor", dataHora);
            command.ExecuteNonQuery();

            int qtdVencer = Convert.ToInt32(command.ExecuteScalar());
            con.Close();

            //verifica se tem boletos a vencer
            if (qtdVencer > 0)
            {
                if (minutos == 57 && segundos == 10 && milisegundos >= 600)
                {
                    //exibe o icone
                    notifyIcon1.Visible = true;
                    //texto a ser exibido da notificação
                    notifyIcon1.Text = "ATENÇÃO";
                    //titulo da mensagem
                    notifyIcon1.BalloonTipTitle = "Boletos a Vencer!";
                    //texto da mensagem
                    if (qtdVencer > 0)
                    {
                        notifyIcon1.BalloonTipText = "Você Possui " + qtdVencer.ToString() + "  boletos à vencer dentro de alguns Dias";
                    }
                    else
                    {
                        notifyIcon1.BalloonTipText = "Você Possui " + qtdVencer.ToString() + " boleto à vencer dentro de alguns Dias";
                    }

                    //o tempo em que ficara sendo exibido
                    notifyIcon1.ShowBalloonTip(1000);
                }
            }
        }
    }
}
