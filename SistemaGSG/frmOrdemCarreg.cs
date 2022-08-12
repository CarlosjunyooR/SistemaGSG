using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class frmOrdemCarreg : MetroFramework.Forms.MetroForm
    {
        public frmOrdemCarreg()
        {
            InitializeComponent();
        }
        DataTable table = new DataTable();
        private void frmOrdemCarreg_Load(object sender, EventArgs e)
        {
            DataGridBancoDados();
        }
        private void DataGridBancoDados()
        {
            table.Columns.Add("_1", typeof(string));
            table.Columns.Add("_2", typeof(string));
            table.Columns.Add("_3", typeof(string));
            table.Columns.Add("_4", typeof(string));
            table.Columns.Add("_5", typeof(string));
            table.Columns.Add("_6", typeof(string));
            table.Columns.Add("_7", typeof(string));
            table.Columns.Add("_8", typeof(string));
            table.Columns.Add("_9", typeof(string));
            table.Columns.Add("_10", typeof(string));
            table.Columns.Add("_11", typeof(string));
            table.Columns.Add("_12", typeof(string));
            table.Columns.Add("_13", typeof(string));
            table.Columns.Add("_14", typeof(string));
            table.Columns.Add("_15", typeof(string));
            table.Columns.Add("_16", typeof(string));
            table.Columns.Add("_17", typeof(string));
            table.Columns.Add("_18", typeof(string));
            table.Columns.Add("_19", typeof(string));
            table.Columns.Add("_20", typeof(string));
            table.Columns.Add("_21", typeof(string));
            table.Columns.Add("_22", typeof(string));
            table.Columns.Add("_23", typeof(string));
            table.Columns.Add("NOME MOTORISTA", typeof(string));
            table.Columns.Add("_25", typeof(string));
            table.Columns.Add("_26", typeof(string));
            DT_SAP.DataSource = table;
        }
        private void ComboBox()
        {
            StatusProgressBar.Value = 100;


            MySqlCommand com = new MySqlCommand();
            com.Connection = ConexaoDados.GetConnectionFaturameto();
            com.CommandText = "SELECT * FROM tb_ordem_motorista";
            MySqlDataReader dr = com.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            ComBoxNomeMotorista.DisplayMember = "nome_motorista";
            ComBoxNomeMotorista.DataSource = dt;
            ConexaoDados.GetConnectionFaturameto().Close();
        }
        private void ConsultaConexao()
        {
            try
            {

                if (ConexaoDados.GetConnectionFaturameto().State == ConnectionState.Open)
                {
                    StatusOnline.Visible = true;
                    StatusOffline.Visible = false;
                }
                ConexaoDados.GetConnectionFaturameto().Close();
            }
            catch (MySqlException)
            {
                StatusOnline.Visible = false;
                StatusOffline.Visible = true;
                Temporizador.Enabled = false;
            }
        }
        private void BaixaSAP()
        {
            try
            {
                //Pega a tela de execução do Windows
                CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
                //Pega a entrada ROT para o SAP Gui para conectar-se ao COM
                object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
                //Pega a referência de Scripting Engine do SAP
                object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
                //Pega a referência da janela de aplicativos em execução no SAP
                GuiApplication GuiApp = (GuiApplication)engine;
                //Pega a primeira conexão aberta do SAP
                GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
                //Pega a primeira sessão aberta
                GuiSession Session = (GuiSession)connection.Children.ElementAt(0);
                //Pega a referência ao "FRAME" principal para enviar comandos de chaves virtuais o SAP
                GuiFrameWindow guiWindow = Session.ActiveWindow;
                //Abre Transação
                StatusProgressBar.Value = 0;
                Session.SendCommand("/NZBL023");
                guiWindow.SendVKey(0);
                StatusProgressBar.Value = 15;
                ((GuiRadioButton)Session.FindById("wnd[0]/usr/radRB_USGA")).Select();
                ((GuiRadioButton)Session.FindById("wnd[0]/usr/radRB_ACESSO")).Select();
                ((GuiGridView)Session.FindById("wnd[0]/usr/cntlZCUSTOMCONTROL/shellcont/shell")).PressToolbarContextButton("&MB_EXPORT");
                ((GuiGridView)Session.FindById("wnd[0]/usr/cntlZCUSTOMCONTROL/shellcont/shell")).SelectContextMenuItem("&PC");
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = @"C:\ArquivosSAP\";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "FATURAMENTO.txt";
                StatusProgressBar.Value = 25;
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).CaretPosition = 6;
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                Session.SendCommand("/N");
                StatusProgressBar.Value = 30;
            }
            catch
            {
                StatusSAP.Visible = true;
                checkBox1.Checked = false;
                MessageBox.Show("O programa SAP encontra-se fechado!\n quando estiver pronto me avise!...\n Marque o box no inicio e clique novamente em Consulta.");
            }
        }
        private void ImportarTXT()
        {
            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\FATURAMENTO.txt");
            string[] values;
            StatusProgressBar.Value = 50;
            for (int i = 5; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('|');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim('-');
                }
                table.Rows.Add(row);
            }
            StatusProgressBar.Value = 55;
        }
        private void PreencherTextBox()
        {
            int countg = DT_SAP.RowCount;
            int numero = 0;
            while (numero < countg)
            {
                try
                {
                    NomeMotorista.Text = DT_SAP.Rows[numero].Cells[23].Value.ToString().Trim();
                    ImportarDataGrid();
                    numero++;
                }
                catch
                {
                    break;
                }
            }
            StatusProgressBar.Value = 75;
        }
        private void ImportarDataGrid()
        {
            try
            {

                if (string.IsNullOrEmpty(NomeMotorista.Text))
                {

                }
                else
                {
                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_motorista (`nome_motorista`) " +
                    "VALUES " +
                    "('" + NomeMotorista.Text + "')", ConexaoDados.GetConnectionFaturameto());
                    cmd.ExecuteNonQuery();
                }
                ConexaoDados.GetConnectionFaturameto().Close();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }
        private void ConsultaMotorista()
        {
            TextBoxVez.Text = ComBoxNomeMotorista.SelectedIndex.ToString();
            int Vez = Convert.ToInt32(TextBoxVez.Text);
            Vez++;
            TextBoxVez.Text = Vez.ToString();
        }
        private void Temporizador_Tick(object sender, EventArgs e)
        {
            ConsultaMotorista();
            ConsultaConexao();
        }
        private void btnConsult_Click(object sender, EventArgs e)
        {

            MySqlCommand comandDell = new MySqlCommand("TRUNCATE tb_ordem_motorista", ConexaoDados.GetConnectionFaturameto());
            comandDell.ExecuteNonQuery();
            ConexaoDados.GetConnectionFaturameto().Close();
            if (checkBox1.Checked)
            {
                BaixaSAP();
                ImportarTXT();
                PreencherTextBox();
                table.Rows.Clear();
                ComboBox();
                Temporizador.Enabled = true;
            }
            else
            {
                Temporizador.Enabled = false;
                TextBoxVez.Text = "Parado.....";
            }
        }
        private void btnPosicao_Click(object sender, EventArgs e)
        {
            frmPosicaoSemana back = new frmPosicaoSemana();
            back.Show();
        }
    }
}
