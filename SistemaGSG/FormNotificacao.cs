using MySql.Data.MySqlClient;
using System;
using System.Windows.Forms;


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

        private void abrirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //para abrir o formulário form1 mesmo, no seu caso, para você seria a tela de login ou o principal, ou outra tela mesmo
            new frm_Main().Show();
        }

        private void contextMenuStrip1_Opening(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void FormNotificacao_Load(object sender, EventArgs e)
        {

        }


        DateTime dataHora;
        DateTime dataHora2 = DateTime.Now;
        Int32 segundos, minutos, milisegundos;

        public int qtdVencer;
        public int qtdVencerp;

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {

                MySqlCommand cmdaa = new MySqlCommand("SELECT CURDATE()", ConexaoDados.GetConnectionEquatorial());
                DateTime DataServ = Convert.ToDateTime(cmdaa.ExecuteScalar());
                string novadata22 = DataServ.AddDays(+15).ToShortDateString();
                DateTime dataHora42 = DateTime.Now;

                MySqlCommand cmd = new MySqlCommand("SELECT CURDATE()", ConexaoDados.GetConnectionEquatorial());
                DateTime DataServidor = Convert.ToDateTime(cmd.ExecuteScalar());
                string novadata = DataServidor.AddDays(+15).ToString("yyyy-MM-dd");

                dataHora42 = DataServidor;
                minutos = dataHora42.Minute;
                segundos = dataHora42.Second;
                milisegundos = dataHora42.Millisecond;

                MySqlCommand command1 = new MySqlCommand("SELECT COUNT(*) FROM notifica_vencimento WHERE data BETWEEN @DataServidor AND @dataFuturo", ConexaoDados.GetConnectionEquatorial());

                command1.Parameters.AddWithValue("@DataServidor", dataHora2.ToString("yyyy-MM-dd"));
                command1.Parameters.AddWithValue("@dataFuturo", novadata);
                command1.ExecuteNonQuery();

                int qtdVencer = Convert.ToInt32(command1.ExecuteScalar());

                ConexaoDados.GetConnectionEquatorial().Close();

                //verifica se tem boletos a vencer
                if (qtdVencer > 0)
                {
                    int i = 0;
                    i++;
                    if (i > 0 && i <= 9)
                    {
                        notifyIcon1.Visible = true;
                        notifyIcon1.Text = "INFORMAÇÃO!";
                        notifyIcon1.BalloonTipTitle = "BOLETOS CEAL & CELPE";

                        if (qtdVencer > 1)
                        {
                            notifyIcon1.BalloonTipText = "Você Possui " + qtdVencer.ToString() + " boletos para vencer no prazo de quinze dias!";
                        }
                        else
                        {
                            notifyIcon1.BalloonTipText = "Você Possui " + qtdVencer.ToString() + " boletos para vencer no prazo de quinze dias!";
                        }
                        notifyIcon1.ShowBalloonTip(1000);
                    }
                }
                else
                {
                    notifyIcon1.BalloonTipText = "Não há boletos pra vencer!";
                }
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
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
    }
}
