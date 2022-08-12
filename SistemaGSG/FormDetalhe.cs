using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace SistemaGSG
{
    public partial class FormDetalhe : MetroFramework.Forms.MetroForm
    {
        public FormDetalhe(string text, string text1, string text2, string text3)
        {
            InitializeComponent();
            txtCodigoFornecedor.Text = text;
            txtDesconto.Text = text1;
            txtSafra.Text = text2;
            txtIDdetalhe.Text = text3;
        }
        private void SAPagamento()
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
                //Maximisa Janela
                guiWindow.Maximize();
                //Abre Transação
                Session.SendCommand("/NZSCREL001");
                //Inicia a Barra de Progresso em 25%
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkP_GRPVAN")).Selected = true;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_WERKS")).Text = "USGA";
                ((GuiComboBox)Session.FindById("wnd[0]/usr/cmbP_SAFRA")).Key = "2020/2021";
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_PERIO")).Text = "36";
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[33]")).Press();
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).SetCurrentCell(6, "TEXT");
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).SelectedRows = "6";
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).ClickCurrentCell();
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[45]")).Press();
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = @"C:\ArquivosSAP\";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "RELFORNECEDOR.txt";
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                //Pega a Barra de Status do SAP
                //GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                //if (statusbar.MessageType == "S")
                //{
                //MessageBox.Show("SUCESSO!");
                //}
                CarregarDataGrid();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }
        DataTable table = new DataTable();
        private void LerDataGrid()
        {
            try
            {
                string contrato;
                string codigo;
                string desconto;

                int cont = 10;

                txtContratoFornecedor.Text = dataGridView1.SelectedRows[cont].Cells[1].Value.ToString();
                txtCodigoFornecedor.Text = dataGridView1.SelectedRows[cont].Cells[2].Value.ToString();
                txtDesconto.Text = dataGridView1.SelectedRows[cont].Cells[5].Value.ToString();

                //MessageBox.Show(contrato.ToString());

            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }

        }
        private void CarregarDataGrid()
        {

            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\RELFORNECEDOR.txt", Encoding.UTF7);
            string[] values;
            for (int i = 6; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('|');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim('-').Trim();
                }
                dataGridView1.Rows.Add(row);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SAPagamento();
        }

        private void FormDetalhe_Load(object sender, EventArgs e)
        {
        }

        private void ricTexto_DoubleClick(object sender, EventArgs e)
        {
            ricTexto.Text = "";
        }

        private void textValor1_DoubleClick(object sender, EventArgs e)
        {
            if (ricTexto.Enabled == false)
            {
                ricTexto.Enabled = true;
            }
            else
            {
                ricTexto.Enabled = false;
                ricTexto.Text = "Digite aqui, Informações referente ao desconto...";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                dataDesconto.Format = DateTimePickerFormat.Custom;
                dataDesconto.CustomFormat = "yyyy-MM-dd";

                if (ricTexto.Text == "Digite aqui, Informações referente ao desconto...")
                {
                    MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_detalhamentodebitos (col_Fornec, col_Contrato, col_TipoDesc, col_Data, col_ValorFixo, col_TextoDesconto, col_dataImport, col_Safra,id_Debito) VALUES " +
                        "('" + txtCodigoFornecedor.Text + "', '" + txtContratoFornecedor.Text + "', '" + txtDesconto.Text + "', '" + this.dataDesconto.Text + "', '" + txtValorDesconto.Text.Replace("R$ ", "") + "', NULL, NOW(), '" + txtSafra.Text + "','" + txtIDdetalhe.Text + "')", ConexaoDados.GetConnectionFornecedor());
                    prompt_cmd.ExecuteNonQuery();
                    txtContratoFornecedor.Text = "";
                    txtValorDesconto.Text = "";
                }
                else
                {
                    MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_detalhamentodebitos (col_Fornec, col_Contrato, col_TipoDesc, col_Data, col_ValorFixo, col_TextoDesconto, col_dataImport, col_Safra,id_Debito) VALUES " +
                        "('" + txtCodigoFornecedor.Text + "', '" + txtContratoFornecedor.Text + "', '" + txtDesconto.Text + "', '" + this.dataDesconto.Text + "', '" + txtValorDesconto.Text.Replace("R$ ", "") + "', '" + ricTexto.Text + "', NOW(), '" + txtSafra.Text + "','" + txtIDdetalhe.Text + "')", ConexaoDados.GetConnectionFornecedor());
                    prompt_cmd.ExecuteNonQuery();
                    txtContratoFornecedor.Text = "";
                    txtValorDesconto.Text = "";
                    ricTexto.Text = "Digite aqui, Informações referente ao desconto...";
                }
                ConexaoDados.GetConnectionFornecedor().Close();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            dataDesconto.Format = DateTimePickerFormat.Custom;
            dataDesconto.CustomFormat = "dd/MM/yyyy";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LerDataGrid();
        }
    }
}
