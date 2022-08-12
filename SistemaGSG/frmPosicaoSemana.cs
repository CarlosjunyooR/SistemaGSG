using MetroFramework;
using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class frmPosicaoSemana : MetroFramework.Forms.MetroForm
    {
        public frmPosicaoSemana()
        {
            InitializeComponent();
            OvAlter = 0;
        }
        int OvAlter = 0;
        DataTable table = new DataTable();
        private void BaixaSAP()
        {
            try
            {
                LblStatus.ForeColor = Color.Chartreuse;
                LblStatus.Text = "Conectando com o SAP.......";
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
                ProgressBar.Value = 0;
                Session.SendCommand("/NZSD014");
                LblStatus.Text = "Conexão bem sucedida.......";
                guiWindow.SendVKey(0);
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_DOCDAT-LOW")).Text = this.date1.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_DOCDAT-HIGH")).Text = this.date2.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_WERKS-LOW")).Text = "USGA";
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS1_VKORG-LOW")).Text = "OVSG";
                LblStatus.Text = "Aguardando o SAP carregar o período.....";
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[33]")).Press();
                ((GuiComboBox)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX")).Key = "X";
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).CurrentCellColumn = "TEXT";
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).SelectedRows = "0";
                ((GuiGridView)Session.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")).ClickCurrentCell();
                LblStatus.Text = "Iniciando o processo de Download.....";
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[45]")).Press();
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                LblStatus.Text = "Selecionando a Pasta onde o Arquivo será Baixado.....";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = @"C:\ArquivosSAP\";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "POSICAOSEMANA.txt";
                LblStatus.Text = "Arquivo POSICAOSEMANA.txt Baixado com sucesso!.....";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).CaretPosition = 6;
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                Session.SendCommand("/N");

                ImportarTXT();
                PreencherTextBox();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                LblStatus.ForeColor = Color.Red;
                LblStatus.Text = "SAP encontra-se fechado.....";
            }
        }
        private void BaixaSAPContrAcucar()
        {
            try
            {
                LblStatus.ForeColor = Color.Chartreuse;
                LblStatus.Text = "Conectando com o SAP.......";
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
                ProgressBar.Value = 0;
                Session.SendCommand("/NZSD023");
                LblStatus.Text = "Conexão bem sucedida.......";
                guiWindow.SendVKey(0);
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_WERKS")).Text = "USGA";
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtSO_ERDT-LOW")).Text = this.date1.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtSO_ERDT-HIGH")).Text = this.date2.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtSO_ERDT-HIGH")).SetFocus();
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtSO_ERDT-HIGH")).CaretPosition = 10;
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[45]")).Press();
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                LblStatus.Text = "Iniciando o processo de Download.....";
                LblStatus.Text = "Selecionando a Pasta onde o Arquivo será Baixado.....";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = "C:\\ArquivosSAP\\";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "CONTRATOACUCAR.txt";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).CaretPosition = 17;
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                Session.SendCommand("/N");
                LerTXTContratoAcucar();
                PreencherTextBoxAcucarContrato();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                LblStatus.ForeColor = Color.Red;
                LblStatus.Text = "SAP encontra-se fechado.....";
            }
        }
        private void BaixaSAPOVAcucar()
        {
            try
            {
                LblStatus.ForeColor = Color.Chartreuse;
                LblStatus.Text = "Conectando com o SAP.......";
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
                ProgressBar.Value = 0;
                Session.SendCommand("/NSDO1");
                LblStatus.Text = "Conexão bem sucedida.......";
                guiWindow.SendVKey(0);
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_ERDAT-LOW")).Text = this.date1.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_ERDAT-HIGH")).Text = this.date2.Text;
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_VKORG-LOW")).Text = "OVSG";
                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtS_SPART-LOW")).Text = "20";
                LblStatus.Text = "Aguardando o SAP carregar o período.....";
                ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                ((GuiMenu)Session.FindById("wnd[0]/mbar/menu[0]/menu[4]/menu[2]")).Select();
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                LblStatus.Text = "Iniciando o processo de Download.....";
                LblStatus.Text = "Selecionando a Pasta onde o Arquivo será Baixado.....";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_PATH")).Text = @"C:\ArquivosSAP\";
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "OVACUCAR.txt";
                LblStatus.Text = "Arquivo POSICAOACUCAR.txt Baixado com sucesso!.....";
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                Session.SendCommand("/N");
                ImportarTXTAcucar();
                PreencherTextBoxAcucar();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                LblStatus.ForeColor = Color.Red;
                LblStatus.Text = "SAP encontra-se fechado.....";
            }
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
            table.Columns.Add("_24", typeof(string));
            table.Columns.Add("_25", typeof(string));
            table.Columns.Add("_26", typeof(string));
            table.Columns.Add("_27", typeof(string));
            table.Columns.Add("_28", typeof(string));
            table.Columns.Add("_29", typeof(string));
            table.Columns.Add("_30", typeof(string));
            table.Columns.Add("_31", typeof(string));
            table.Columns.Add("_32", typeof(string));
            table.Columns.Add("_33", typeof(string));
            table.Columns.Add("_34", typeof(string));
            table.Columns.Add("_35", typeof(string));
            table.Columns.Add("_36", typeof(string));
            table.Columns.Add("_37", typeof(string));
            table.Columns.Add("_38", typeof(string));
            table.Columns.Add("_39", typeof(string));
            table.Columns.Add("_40", typeof(string));
            table.Columns.Add("_41", typeof(string));
            table.Columns.Add("_42", typeof(string));
            table.Columns.Add("_43", typeof(string));
            table.Columns.Add("_44", typeof(string));
            table.Columns.Add("_45", typeof(string));
            table.Columns.Add("_46", typeof(string));
            table.Columns.Add("_47", typeof(string));
            table.Columns.Add("_48", typeof(string));
            table.Columns.Add("_49", typeof(string));
            table.Columns.Add("_50", typeof(string));
            table.Columns.Add("_51", typeof(string));
            table.Columns.Add("_52", typeof(string));
            table.Columns.Add("_53", typeof(string));
            table.Columns.Add("_54", typeof(string));
            table.Columns.Add("_55", typeof(string));
            table.Columns.Add("_56", typeof(string));
            table.Columns.Add("_57", typeof(string));
            table.Columns.Add("_58", typeof(string));
            table.Columns.Add("_59", typeof(string));
            table.Columns.Add("_60", typeof(string));
            table.Columns.Add("_61", typeof(string));
            table.Columns.Add("_62", typeof(string));
            table.Columns.Add("_63", typeof(string));
            table.Columns.Add("_64", typeof(string));
            table.Columns.Add("_65", typeof(string));
            table.Columns.Add("_66", typeof(string));
            table.Columns.Add("_67", typeof(string));
            table.Columns.Add("_68", typeof(string));
            table.Columns.Add("_69", typeof(string));
            table.Columns.Add("_70", typeof(string));
            table.Columns.Add("_71", typeof(string));
            table.Columns.Add("_72", typeof(string));
            table.Columns.Add("_73", typeof(string));
            table.Columns.Add("_74", typeof(string));
            table.Columns.Add("_75", typeof(string));
            table.Columns.Add("_76", typeof(string));
            table.Columns.Add("_77", typeof(string));
            table.Columns.Add("_78", typeof(string));
            table.Columns.Add("_79", typeof(string));
            table.Columns.Add("_80", typeof(string));
            table.Columns.Add("_81", typeof(string));
            table.Columns.Add("_82", typeof(string));
            table.Columns.Add("_83", typeof(string));
            table.Columns.Add("_84", typeof(string));
            table.Columns.Add("_85", typeof(string));
            table.Columns.Add("_86", typeof(string));
            table.Columns.Add("_87", typeof(string));
            table.Columns.Add("_88", typeof(string));
            table.Columns.Add("_89", typeof(string));
            table.Columns.Add("_90", typeof(string));
            table.Columns.Add("_91", typeof(string));
            table.Columns.Add("_92", typeof(string));
            table.Columns.Add("_93", typeof(string));
            table.Columns.Add("_94", typeof(string));
            table.Columns.Add("_95", typeof(string));
            table.Columns.Add("_96", typeof(string));
            table.Columns.Add("_97", typeof(string));
            table.Columns.Add("_98", typeof(string));
            table.Columns.Add("_99", typeof(string));
            table.Columns.Add("_100", typeof(string));
            table.Columns.Add("_101", typeof(string));
            table.Columns.Add("_102", typeof(string));
            table.Columns.Add("_103", typeof(string));
            table.Columns.Add("_104", typeof(string));
            table.Columns.Add("_105", typeof(string));
            table.Columns.Add("_106", typeof(string));
            table.Columns.Add("_107", typeof(string));
            table.Columns.Add("_108", typeof(string));
            table.Columns.Add("_109", typeof(string));
            table.Columns.Add("_110", typeof(string));
            DT_SAP.DataSource = table;
        }
        private void DeleteBetween()
        {
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "yyyy-MM-dd";
            date2.Format = DateTimePickerFormat.Custom;
            date2.CustomFormat = "yyyy-MM-dd";
            MySqlCommand comandDell = new MySqlCommand("DELETE FROM tb_saida_semana WHERE DATA_EMISS BETWEEN '" + this.date1.Text + "' AND '" + this.date2.Text + "'", ConexaoDados.GetConnectionFaturameto());
            comandDell.ExecuteNonQuery();
            MySqlCommand comandDelete = new MySqlCommand("DELETE FROM tb_filaphp WHERE dataPeriodo BETWEEN '" + this.date1.Text + "' AND '" + this.date2.Text + "'", ConexaoDados.GetConnectionFaturameto());
            comandDelete.ExecuteNonQuery();
        }
        private void DeleteData()
        {
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "yyyy-MM-dd";
            MySqlCommand comandDell = new MySqlCommand("DELETE FROM tb_boletim WHERE datadoDia = '" + this.date1.Text + "'", ConexaoDados.GetConnectionFaturameto());
            comandDell.ExecuteNonQuery();
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "dd/MM/yyyy";
        }
        private void ImportarTXT()
        {
            LblStatus.Text = "Importando todo arquivo no Banco de Dados, está quase tudo pronto.....";

            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\POSICAOSEMANA.txt", Encoding.UTF7);
            string[] values;
            for (int i = 10; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('|');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim('-');
                }
                table.Rows.Add(row);
            }
        }
        private void ImportarTXTAcucar()
        {
            LblStatus.Text = "Importando todo arquivo no Banco de Dados, está quase tudo pronto.....";

            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\OVACUCAR.txt", Encoding.UTF7);
            string[] values;
            for (int i = 6; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('|');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim('-');
                }
                table.Rows.Add(row);
            }
        }
        private void LerTXTContratoAcucar()
        {
            LblStatus.Text = "Importando todo arquivo no Banco de Dados, está quase tudo pronto.....";

            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\CONTRATOACUCAR.txt", Encoding.UTF7);
            string[] values;
            for (int i = 6; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('|');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim('-');
                }
                table.Rows.Add(row);
            }
        }
        private void PreencherTextBox()
        {
            _28.Format = DateTimePickerFormat.Custom;
            _28.CustomFormat = "yyyy-MM-dd";
            date2.Format = DateTimePickerFormat.Custom;
            date2.CustomFormat = "yyyy-MM-dd";

            int countg = DT_SAP.RowCount;
            int numero = 0;
            int Progresso = 0;
            while (numero < countg)
            {
                try
                {
                    ProgressBar.Value = Progresso;
                    _1.Text = DT_SAP.Rows[numero].Cells[1].Value.ToString().Trim();
                    _2.Text = DT_SAP.Rows[numero].Cells[2].Value.ToString().Trim();
                    _3.Text = DT_SAP.Rows[numero].Cells[3].Value.ToString().Trim();
                    _4.Text = DT_SAP.Rows[numero].Cells[4].Value.ToString().Trim();
                    _5.Text = DT_SAP.Rows[numero].Cells[5].Value.ToString().Trim();
                    _6.Text = DT_SAP.Rows[numero].Cells[6].Value.ToString().Trim();
                    _7.Text = DT_SAP.Rows[numero].Cells[7].Value.ToString().Trim();
                    _8.Text = DT_SAP.Rows[numero].Cells[8].Value.ToString().Trim();
                    _9.Text = DT_SAP.Rows[numero].Cells[9].Value.ToString().Trim();
                    _10.Text = DT_SAP.Rows[numero].Cells[10].Value.ToString().Trim();
                    _11.Text = DT_SAP.Rows[numero].Cells[11].Value.ToString().Trim();
                    _12.Text = DT_SAP.Rows[numero].Cells[12].Value.ToString().Trim();
                    _13.Text = DT_SAP.Rows[numero].Cells[13].Value.ToString().Trim();
                    _14.Text = DT_SAP.Rows[numero].Cells[14].Value.ToString().Trim();
                    _15.Text = DT_SAP.Rows[numero].Cells[15].Value.ToString().Trim();
                    _16.Text = DT_SAP.Rows[numero].Cells[16].Value.ToString().Trim();
                    _17.Text = DT_SAP.Rows[numero].Cells[17].Value.ToString().Trim();
                    _18.Text = DT_SAP.Rows[numero].Cells[18].Value.ToString().Trim();
                    _19.Text = DT_SAP.Rows[numero].Cells[19].Value.ToString().Trim();
                    _20.Text = DT_SAP.Rows[numero].Cells[20].Value.ToString().Trim();
                    _21.Text = DT_SAP.Rows[numero].Cells[21].Value.ToString().Trim();
                    _22.Text = DT_SAP.Rows[numero].Cells[22].Value.ToString().Trim();
                    _23.Text = DT_SAP.Rows[numero].Cells[23].Value.ToString().Trim();
                    _24.Text = DT_SAP.Rows[numero].Cells[24].Value.ToString().Trim();
                    _25.Text = DT_SAP.Rows[numero].Cells[25].Value.ToString().Trim();
                    _26.Text = DT_SAP.Rows[numero].Cells[26].Value.ToString().Trim();
                    _27.Text = DT_SAP.Rows[numero].Cells[27].Value.ToString().Trim();
                    this._28.Text = DT_SAP.Rows[numero].Cells[28].Value.ToString().Replace(".", "/").Trim();
                    _29.Text = DT_SAP.Rows[numero].Cells[29].Value.ToString().Trim();
                    _30.Text = DT_SAP.Rows[numero].Cells[30].Value.ToString().Trim();
                    _31.Text = DT_SAP.Rows[numero].Cells[31].Value.ToString().Trim();
                    _32.Text = DT_SAP.Rows[numero].Cells[32].Value.ToString().Trim();
                    _33.Text = DT_SAP.Rows[numero].Cells[33].Value.ToString().Trim();
                    _34.Text = DT_SAP.Rows[numero].Cells[34].Value.ToString().Trim();
                    _35.Text = DT_SAP.Rows[numero].Cells[35].Value.ToString().Trim();
                    _36.Text = DT_SAP.Rows[numero].Cells[36].Value.ToString().Trim();
                    _37.Text = DT_SAP.Rows[numero].Cells[37].Value.ToString().Trim();
                    _38.Text = DT_SAP.Rows[numero].Cells[38].Value.ToString().Trim();
                    _39.Text = DT_SAP.Rows[numero].Cells[39].Value.ToString().Trim();
                    _40.Text = DT_SAP.Rows[numero].Cells[40].Value.ToString().Trim();
                    _41.Text = DT_SAP.Rows[numero].Cells[41].Value.ToString().Trim();
                    _42.Text = DT_SAP.Rows[numero].Cells[42].Value.ToString().Trim();
                    _43.Text = DT_SAP.Rows[numero].Cells[43].Value.ToString().Trim();
                    _44.Text = DT_SAP.Rows[numero].Cells[44].Value.ToString().Trim();
                    _45.Text = DT_SAP.Rows[numero].Cells[45].Value.ToString().Trim();
                    _46.Text = DT_SAP.Rows[numero].Cells[46].Value.ToString().Trim();
                    _47.Text = DT_SAP.Rows[numero].Cells[47].Value.ToString().Trim();
                    _48.Text = DT_SAP.Rows[numero].Cells[48].Value.ToString().Trim();
                    _49.Text = DT_SAP.Rows[numero].Cells[49].Value.ToString().Trim();
                    _50.Text = DT_SAP.Rows[numero].Cells[50].Value.ToString().Trim();
                    _51.Text = DT_SAP.Rows[numero].Cells[51].Value.ToString().Trim();
                    _52.Text = DT_SAP.Rows[numero].Cells[52].Value.ToString().Trim();
                    _53.Text = DT_SAP.Rows[numero].Cells[53].Value.ToString().Trim();
                    _54.Text = DT_SAP.Rows[numero].Cells[54].Value.ToString().Trim();
                    _55.Text = DT_SAP.Rows[numero].Cells[55].Value.ToString().Trim();
                    _56.Text = DT_SAP.Rows[numero].Cells[56].Value.ToString().Trim();
                    _57.Text = DT_SAP.Rows[numero].Cells[57].Value.ToString().Trim();
                    _58.Text = DT_SAP.Rows[numero].Cells[58].Value.ToString().Trim();
                    _59.Text = DT_SAP.Rows[numero].Cells[59].Value.ToString().Trim();
                    _60.Text = DT_SAP.Rows[numero].Cells[60].Value.ToString().Trim();
                    _61.Text = DT_SAP.Rows[numero].Cells[61].Value.ToString().Trim();
                    _62.Text = DT_SAP.Rows[numero].Cells[62].Value.ToString().Trim();
                    _63.Text = DT_SAP.Rows[numero].Cells[63].Value.ToString().Trim();
                    _64.Text = DT_SAP.Rows[numero].Cells[64].Value.ToString().Trim();
                    _65.Text = DT_SAP.Rows[numero].Cells[65].Value.ToString().Trim();
                    _66.Text = DT_SAP.Rows[numero].Cells[66].Value.ToString().Trim();
                    _67.Text = DT_SAP.Rows[numero].Cells[67].Value.ToString().Trim();
                    _68.Text = DT_SAP.Rows[numero].Cells[68].Value.ToString().Trim();
                    _69.Text = DT_SAP.Rows[numero].Cells[69].Value.ToString().Trim();
                    _70.Text = DT_SAP.Rows[numero].Cells[70].Value.ToString().Trim();
                    _71.Text = DT_SAP.Rows[numero].Cells[71].Value.ToString().Trim();
                    _72.Text = DT_SAP.Rows[numero].Cells[72].Value.ToString().Trim();
                    _73.Text = DT_SAP.Rows[numero].Cells[73].Value.ToString().Trim();
                    _74.Text = DT_SAP.Rows[numero].Cells[74].Value.ToString().Trim();
                    _75.Text = DT_SAP.Rows[numero].Cells[75].Value.ToString().Trim();
                    _76.Text = DT_SAP.Rows[numero].Cells[76].Value.ToString().Trim();
                    //_77.Text = DT_SAP.Rows[numero].Cells[77].Value.ToString().Trim();
                    ImportarDataGrid();
                    numero++;
                    Progresso++;
                }
                catch (Exception ErroProg)
                {
                    MessageBox.Show(ErroProg.Message);
                    LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
                    break;
                }
            }
            ProgressBar.Value = 1000;
            LblStatus.Text = "Pronto processo finalizado.....";
        }
        private void PreencherTextBoxAcucarContrato()
        {
            _28.Format = DateTimePickerFormat.Custom;
            _28.CustomFormat = "yyyy-MM-dd";
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "yyyy-MM-dd";
            date2.Format = DateTimePickerFormat.Custom;
            date2.CustomFormat = "yyyy-MM-dd";
            int countg = DT_SAP.RowCount;
            int numero = 0;
            int Progresso = 0;
            while (numero < countg)
            {
                try
                {
                    ProgressBar.Value = Progresso;
                    string Contrato = DT_SAP.Rows[numero].Cells[2].Value.ToString().Trim();
                    string Itm = DT_SAP.Rows[numero].Cells[3].Value.ToString().Trim();
                    string TpDV = DT_SAP.Rows[numero].Cells[4].Value.ToString().Trim();
                    string Cen = DT_SAP.Rows[numero].Cells[5].Value.ToString().Trim();
                    string CodCliente = DT_SAP.Rows[numero].Cells[6].Value.ToString().Trim();
                    string NomeCliente = DT_SAP.Rows[numero].Cells[7].Value.ToString().Trim();
                    string Material = DT_SAP.Rows[numero].Cells[8].Value.ToString().Trim();
                    string Denominação = DT_SAP.Rows[numero].Cells[9].Value.ToString().Trim();
                    string NoContCl = DT_SAP.Rows[numero].Cells[10].Value.ToString().Trim();
                    string QtdCont = DT_SAP.Rows[numero].Cells[11].Value.ToString().Trim();
                    string QtdPedida = DT_SAP.Rows[numero].Cells[12].Value.ToString().Trim();
                    string Devolução = DT_SAP.Rows[numero].Cells[13].Value.ToString().Trim();
                    string QtdPend = DT_SAP.Rows[numero].Cells[14].Value.ToString().Trim();
                    string UMB = DT_SAP.Rows[numero].Cells[15].Value.ToString().Trim();
                    if (string.IsNullOrEmpty(Itm))
                    {

                    }
                    else
                    {
                        if (string.IsNullOrEmpty(TpDV))
                        {

                        }
                        else
                        {
                            MySqlCommand com = new MySqlCommand();
                            com.Connection = ConexaoDados.GetConnectionFaturameto();
                            com.CommandText = "SELECT * FROM tb_contratos WHERE col_docvendas='" + Contrato + "'";
                            MySqlDataReader dr = com.ExecuteReader();
                            DataTable dt = new DataTable();
                            dt.Load(dr);
                            //Contagem de Linhas
                            int ContratosTB = dt.Rows.Count;
                            if (ContratosTB == 0)
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO `tb_contratos` (`col_docvendas`, `col_item`, `col_tpdv`, `col_centro`, `col_cliente`, `col_nome`, `col_material`, `SAFRA`, `col_nocontrato_cliente`, `col_qtd_contrato`, `col_qtd_pedida`, `col_devolucao`, `col_qtd_pendente`, `col_un`, `col_data`, `col_data_hora_import`, `col_position`)" +
                            "VALUES " +
                            "('" + Contrato + "','" + Itm + "','" + TpDV + "','" + Cen + "','" + CodCliente + "','" + NomeCliente + "','" + Material + "','" + comboBox1.Text + "', '" + NoContCl + "', '" + QtdCont.Replace(".", "").Replace(",", ".") + "', '" + QtdPedida.Replace(".", "").Replace(",", ".") + "', '" + Devolução.Replace(".", "").Replace(",", ".") + "', '" + QtdPend.Replace(".", "").Replace(",", ".") + "', '" + UMB + "', '" + date1.Text + "', NOW(), '1')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    numero++;
                    Progresso++;
                }
                catch (Exception ErroProg)
                {
                    LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
                    MessageBox.Show(ErroProg.Message);
                    break;
                }
            }
            LimparTEXT();
            ProgressBar.Value = 1000;
            LblStatus.Text = "Pronto processo finalizado.....";
            ConexaoDados.GetConnectionFaturameto().Close();
        }
        private void PreencherTextBoxAcucar()
        {
            _28.Format = DateTimePickerFormat.Custom;
            _28.CustomFormat = "yyyy-MM-dd";
            date2.Format = DateTimePickerFormat.Custom;
            date2.CustomFormat = "yyyy-MM-dd";
            int countg = DT_SAP.RowCount;
            int numero = 0;
            int Progresso = 0;
            while (numero < countg)
            {
                try
                {
                    ProgressBar.Value = Progresso;
                    string DocOrdem = DT_SAP.Rows[numero].Cells[1].Value.ToString().Trim();
                    string Itm = DT_SAP.Rows[numero].Cells[2].Value.ToString().Trim();
                    string Div = DT_SAP.Rows[numero].Cells[3].Value.ToString().Trim();
                    string Denominacao = DT_SAP.Rows[numero].Cells[5].Value.ToString().Trim();
                    string TpDv = DT_SAP.Rows[numero].Cells[6].Value.ToString().Trim();
                    this._28.Text = DT_SAP.Rows[numero].Cells[7].Value.ToString().Replace(".", "/").Trim();
                    string QtdConf = DT_SAP.Rows[numero].Cells[8].Value.ToString().Trim();
                    string Pedido = DT_SAP.Rows[numero].Cells[9].Value.ToString().Trim();
                    string MATERIAL = DT_SAP.Rows[numero].Cells[22].Value.ToString().Trim();
                    string UMB = DT_SAP.Rows[numero].Cells[23].Value.ToString().Trim();
                    string DesCliente = DT_SAP.Rows[numero].Cells[24].Value.ToString().Trim();

                    if (string.IsNullOrEmpty(Itm))
                    {

                    }
                    else
                    {
                        if (string.IsNullOrEmpty(Div))
                        {

                        }
                        else
                        {
                            MySqlCommand com = new MySqlCommand();
                            com.Connection = ConexaoDados.GetConnectionFaturameto();
                            com.CommandText = "SELECT * FROM tb_ordem_venda WHERE Doc_SD='" + DocOrdem + "'";
                            MySqlDataReader dr = com.ExecuteReader();
                            DataTable dt = new DataTable();
                            dt.Load(dr);
                            //Contagem de Linhas
                            int ItemOV = dt.Rows.Count;
                            if (ItemOV == 0)
                            {
                                if (MATERIAL == "100000")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100001")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100141")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100002")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100035")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100145")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                                if (MATERIAL == "100180")
                                {
                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                    "VALUES " +
                                    "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '1', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandTimeout = 120; //default 30 segundos
                                }
                            }
                            if (ItemOV == 1)
                            {
                                if (Div == "2")
                                {
                                    if (MATERIAL == "100000")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "', '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100001")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100141")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100002")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100035")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100145")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                    if (MATERIAL == "100180")
                                    {
                                        MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                                        "VALUES " +
                                        "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".", "").Replace(",", ".") + "','" + Pedido + "',  '2', NULL, NULL,  '" + UMB + "', '" + DesCliente + "', NULL,NULL, '4', NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                                        cmd.ExecuteNonQuery();
                                        cmd.CommandTimeout = 120; //default 30 segundos
                                    }
                                }
                            }
                        }
                    }
                    numero++;
                    Progresso++;
                }
                catch (Exception ErroProg)
                {
                    MessageBox.Show(ErroProg.Message);
                    LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
                    break;
                }
            }
            LimparTEXT();
            ProgressBar.Value = 1000;
            LblStatus.Text = "Pronto processo finalizado.....";
        }
        private void ImportarDataGrid()
        {
            try
            {
                if (string.IsNullOrEmpty(_1.Text))
                {

                }
                else
                {
                    _16.Text.Trim();
                    if (_30.Text == "100000")
                    {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                    }
                    if (_30.Text == "100001")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100002")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100014")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100015")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100035")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100141")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100145")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();

                    }
                    if (_30.Text == "100180")
                    {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + comboBox1.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message + " Erro Aqui! Importar tbsemana");

                //MessageBox.Show("Erro com sua conexão, Verifique se o servidor está Online...");
            }
        }
        private void LimparTEXT()
        {
            _1.Text = String.Empty;
            _2.Text = String.Empty;
            _3.Text = String.Empty;
            _4.Text = String.Empty;
            _5.Text = String.Empty;
            _6.Text = String.Empty;
            _7.Text = String.Empty;
            _8.Text = String.Empty;
            _9.Text = String.Empty;
            _10.Text = String.Empty;
            _11.Text = String.Empty;
            _12.Text = String.Empty;
            _13.Text = String.Empty;
            _14.Text = String.Empty;
            _15.Text = String.Empty;
            _16.Text = String.Empty;
            _17.Text = String.Empty;
            _18.Text = String.Empty;
            _19.Text = String.Empty;
            _20.Text = String.Empty;
            _21.Text = String.Empty;
            _22.Text = String.Empty;
            _23.Text = String.Empty;
            _24.Text = String.Empty;
            _25.Text = String.Empty;
            _26.Text = String.Empty;
            _27.Text = String.Empty;
            _29.Text = String.Empty;
            _30.Text = String.Empty;
            _31.Text = String.Empty;
            _32.Text = String.Empty;
            _33.Text = String.Empty;
            _34.Text = String.Empty;
            _35.Text = String.Empty;
            _36.Text = String.Empty;
            _37.Text = String.Empty;
            _38.Text = String.Empty;
            _39.Text = String.Empty;
            _40.Text = String.Empty;
            _41.Text = String.Empty;
            _42.Text = String.Empty;
            _43.Text = String.Empty;
            _44.Text = String.Empty;
            _45.Text = String.Empty;
            _46.Text = String.Empty;
            _47.Text = String.Empty;
            _48.Text = String.Empty;
            _49.Text = String.Empty;
            _50.Text = String.Empty;
            _51.Text = String.Empty;
            _52.Text = String.Empty;
            _53.Text = String.Empty;
            _54.Text = String.Empty;
            _55.Text = String.Empty;
            _56.Text = String.Empty;
            _57.Text = String.Empty;
            _58.Text = String.Empty;
            _59.Text = String.Empty;
            _60.Text = String.Empty;
            _61.Text = String.Empty;
            _62.Text = String.Empty;
            _63.Text = String.Empty;
            _64.Text = String.Empty;
            _65.Text = String.Empty;
            _66.Text = String.Empty;
            _67.Text = String.Empty;
            _68.Text = String.Empty;
            _69.Text = String.Empty;
            _70.Text = String.Empty;
            _71.Text = String.Empty;
            _72.Text = String.Empty;
            _73.Text = String.Empty;
            _74.Text = String.Empty;
            _75.Text = String.Empty;
            _76.Text = String.Empty;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var sw = new Stopwatch();
            sw.Start();

            date1.Text = monthCalendar1.SelectionRange.Start.ToString();
            date2.Text = monthCalendar1.SelectionRange.End.ToString();
            if (RDCompleto.Checked == true)
            {
                DeleteBetween();

                table.Rows.Clear();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd.MM.yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd.MM.yyyy";

                BaixaSAP();
                LimparTEXT();
                ViagensConsultaSAP();

                table.Rows.Clear();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd.MM.yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd.MM.yyyy";
                BaixaSAPContrAcucar();

                table.Rows.Clear();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd.MM.yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd.MM.yyyy";

                BaixaSAPOVAcucar();
                table.Rows.Clear();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd.MM.yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd.MM.yyyy";

                OVSeparar();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd/MM/yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd/MM/yyyy";
            }
            try
            {
                MySqlCommand CommandoMySql = new MySqlCommand("INSERT INTO tb_periodo (`col_inicio`, `col_fim`, `col_ID`) " +
                "VALUES " +
                "('" + this.monthCalendar1.SelectionStart.ToString("yyyy-MM-dd") + "','" + this.monthCalendar1.SelectionEnd.ToString("yyyy-MM-dd") + "','1')", ConexaoDados.GetConnectionFaturameto());
                CommandoMySql.ExecuteNonQuery();
            }
            catch (MySqlException ErroMysql)
            {
                MessageBox.Show("Erro No Banco de Dados!" + ErroMysql, "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ErroCatch)
            {
                MessageBox.Show("Erro No Aplicativo!  " + ErroCatch.Message, "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            ConexaoDados.GetConnectionFaturameto().Close();
            sw.Stop();
            MessageBox.Show("Tempo Decorrido... " + sw.Elapsed.ToString(@"hh\:mm\:ss"));
        }
        private void OVSeparar()
        {
            try
            {
                LblStatus.ForeColor = Color.Chartreuse;
                LblStatus.Text = "Conectando com o SAP.......";
                //System.Diagnostics.Debugger.Break();
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
                ProgressBar.Value = 0;
                if (OvAlter == 0)
                {
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_ordem_venda` WHERE col_TipoOV IS NULL", ConexaoDados.GetConnectionFaturameto());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    DT_SAP.DataSource = SS;
                    ConexaoDados.GetConnectionFaturameto().Close();

                }
                if (OvAlter == 1)
                {
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_ordem_venda` WHERE Div_ITEM='2' AND Qtd_ordem IS NULL", ConexaoDados.GetConnectionFaturameto());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    DT_SAP.DataSource = SS;
                    ConexaoDados.GetConnectionFaturameto().Close();

                }
                dataovalterada1.Format = DateTimePickerFormat.Custom;
                dataovalterada1.CustomFormat = "yyyy-MM-dd";
                dataovalterada2.Format = DateTimePickerFormat.Custom;
                dataovalterada2.CustomFormat = "yyyy-MM-dd";

                int RowCount = DT_SAP.RowCount;
                int EndCount = 0;
                if (RowCount > 0)
                {
                    while (EndCount <= RowCount)
                    {

                        Session.SendCommand("/NVA03");
                        LblStatus.Text = "Conexão bem sucedida.......";
                        ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtVBAK-VBELN")).Text = DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString();
                        ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtVBAK-VBELN")).CaretPosition = 5;
                        guiWindow.SendVKey(0);
                        String GrupoCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtRV45A-TXT_VBELN")).Text;
                        ((GuiTab)Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\05")).Select();
                        String DataOvAlt = ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\05/ssubSUBSCREEN_BODY:SAPMV45A:4405/subSUBSCREEN_TC:SAPMV45A:4920/tblSAPMV45ATCTRL_UEIN_BESCHAFFUNG/ctxtVBEP-MBDAT[5,0]")).Text;
                        dataovalterada1.Text = DataOvAlt;
                        String DataOvAlt2 = ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\05/ssubSUBSCREEN_BODY:SAPMV45A:4405/subSUBSCREEN_TC:SAPMV45A:4920/tblSAPMV45ATCTRL_UEIN_BESCHAFFUNG/txtRV45A-ETDAT[11,0]")).Text;
                        dataovalterada2.Text = DataOvAlt2;

                        Session.SendCommand("/NVA03");
                        if (GrupoCliente == "Venda Entrega Futura")
                        {
                            MySqlCommand cmd = new MySqlCommand("UPDATE `tb_ordem_venda` SET `col_TipoOV` = 'V. Mercado Interno' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "';", ConexaoDados.GetConnectionFaturameto());
                            cmd.ExecuteNonQuery();
                            cmd.CommandTimeout = 120;
                            if (OvAlter == 1)
                            {
                                MySqlCommand ovAlterCommand = new MySqlCommand("UPDATE `tb_ordem_venda` SET `Data_doc` = '" + this.dataovalterada1.Text + "', `Qtd_ordem` = '1' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "' AND `tb_ordem_venda`.`Div_ITEM`='2';", ConexaoDados.GetConnectionFaturameto());
                                ovAlterCommand.ExecuteNonQuery();
                                ovAlterCommand.CommandTimeout = 120;
                                MySqlCommand ovAlterCommand2 = new MySqlCommand("UPDATE `tb_ordem_venda` SET `Data_doc` = '" + this.dataovalterada2.Text + "' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "' AND `tb_ordem_venda`.`Div_ITEM`='1';", ConexaoDados.GetConnectionFaturameto());
                                ovAlterCommand2.ExecuteNonQuery();
                                ovAlterCommand2.CommandTimeout = 120;
                                ConexaoDados.GetConnectionFaturameto().Close();

                            }
                        }
                        else
                        {
                            MySqlCommand cmd = new MySqlCommand("UPDATE `tb_ordem_venda` SET `col_TipoOV` = '" + GrupoCliente.Trim() + "' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "';", ConexaoDados.GetConnectionFaturameto());
                            cmd.ExecuteNonQuery();
                            cmd.CommandTimeout = 120;
                            if (OvAlter == 1)
                            {
                                MySqlCommand ovAlterCommand = new MySqlCommand("UPDATE `tb_ordem_venda` SET `Data_doc` = '" + this.dataovalterada1.Text + "', `Qtd_ordem` = '1' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "' AND `tb_ordem_venda`.`Div_ITEM`='2';", ConexaoDados.GetConnectionFaturameto());
                                ovAlterCommand.ExecuteNonQuery();
                                ovAlterCommand.CommandTimeout = 120;
                                MySqlCommand ovAlterCommand2 = new MySqlCommand("UPDATE `tb_ordem_venda` SET `Data_doc` = '" + this.dataovalterada2.Text + "' WHERE `tb_ordem_venda`.`Doc_SD` = '" + DT_SAP.Rows[EndCount].Cells["Doc_SD"].Value.ToString() + "' AND `tb_ordem_venda`.`Div_ITEM`='1';", ConexaoDados.GetConnectionFaturameto());
                                ovAlterCommand2.ExecuteNonQuery();
                                ovAlterCommand2.CommandTimeout = 120;
                                ConexaoDados.GetConnectionFaturameto().Close();
                            }
                        }
                        EndCount++;
                    }
                }
                Session.SendCommand("/N");
            }
            catch (ArgumentOutOfRangeException)
            {
                LblStatus.Text = "Fim do Processo!.....";
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
                LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
            }
        }
        private void frmPosicaoSemana_Load(object sender, EventArgs e)
        {
            DataGridBancoDados();
            date1.Visible = false;
            date2.Visible = false;
            UserTXT.Text = dados.Usuario;
            try
            {
                MySqlCommand MyCommand = new MySqlCommand();
                MyCommand.Connection = ConexaoDados.GetConnectionFaturameto();
                MyCommand.CommandText = "SELECT * FROM `tb_periodo` ORDER BY `tb_periodo`.`id` DESC";
                MySqlDataReader dreader = MyCommand.ExecuteReader();
                while (dreader.Read())
                {
                    this.label1.Text = dreader["col_inicio"].ToString();
                    this.label3.Text = dreader["col_fim"].ToString();
                    break;
                }
                date1.Text = label1.Text;
                date2.Text = label3.Text;
                date1.Format = DateTimePickerFormat.Custom;
                date2.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd/MM/yyyy";
                date2.CustomFormat = "dd/MM/yyyy";
                monthCalendar1.SelectionStart = date1.Value;
                monthCalendar1.SelectionEnd = date2.Value;
            }
            catch (Exception Error)
            {
                MessageBox.Show(Error.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionFaturameto().Close();
            }
            try
            {
                MySqlCommand com = new MySqlCommand();
                com.Connection = ConexaoDados.GetConnectionPosto();
                com.CommandText = "SELECT * FROM `tb_safra` WHERE `STATUS` != 2";
                MySqlDataReader dr = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                comboBox1.DisplayMember = "id_safra";
                comboBox1.DataSource = dt;
                ConexaoDados.GetConnectionFaturameto().Close();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            finally
            {
                ConexaoDados.GetConnectionPosto().Close();
            }
        }
        private void ViagensConsultaSAP()
        {
            try
            {
                LblStatus.ForeColor = Color.Chartreuse;
                LblStatus.Text = "Baixando Viagens.....";
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
                Session.SendCommand("/NZBL023");
                guiWindow.SendVKey(0);

                ((GuiRadioButton)Session.FindById("wnd[0]/usr/radRB_USGA")).Select();
                _1.Text = ((GuiTextField)Session.FindById("wnd[0]/usr/txtWC_CARRE_INI")).Text;
                _2.Text = ((GuiTextField)Session.FindById("wnd[0]/usr/txtWC_BALA")).Text;
                _3.Text = ((GuiTextField)Session.FindById("wnd[0]/usr/txtWC_PORT")).Text;
                _4.Text = ((GuiTextField)Session.FindById("wnd[0]/usr/txtWC_EXPE")).Text;
                ImportarCarregamento();
                Session.SendCommand("/N");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                LblStatus.Text = "SAP encontra-se fechado.....";
            }
            LblStatus.Text = "Processo Finalizado!.....";
        }
        private void ImportarCarregamento()
        {
            try
            {
                if (string.IsNullOrEmpty(_1.Text))
                {

                }
                else
                {
                    MySqlCommand comandDell = new MySqlCommand("TRUNCATE tb_filasap", ConexaoDados.GetConnectionFaturameto());
                    comandDell.ExecuteNonQuery();

                    MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_filasap (`CARRE_INI`, `BALA`, `PORT`, `EXPE`, `data`, `hora`, `status`) " +
                    "VALUES " +
                    "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "', CURDATE(), NOW(), '3')", ConexaoDados.GetConnectionFaturameto());
                    cmd.ExecuteNonQuery();
                }
                //MetroMessageBox.Show(this, "Processo Finalizado.", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MetroMessageBox.Show(this, "Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main frm_Main = new frm_Main();
                frm_Main.Show();
                Close();
            }
        }
        private void BaixarBoletim()
        {
            LblStatus.ForeColor = Color.Chartreuse;
            LblStatus.Text = "Baixando Boletim.....";
            //Pega a tela de execução do Windows
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            //Pega a entrada ROT para o SAP Gui para conectar-se ao COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Pega a referência de Scripting Engine do SAP
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Pega a referência da janela de aplicativos em execução no SAP
            GuiApplication GuiApp = (GuiApplication)engine;
            //Pega a primeira conexão aberta do SAP
            GuiConnection connectionSAP = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //Pega a primeira sessão aberta
            GuiSession Session = (GuiSession)connectionSAP.Children.ElementAt(0);
            //Pega a referência ao "FRAME" principal para enviar comandos de chaves virtuais o SAP
            GuiFrameWindow guiWindow = Session.ActiveWindow;
            //Abre Transação
            Session.SendCommand("/NZQMBOL");
            guiWindow.SendVKey(0);
            LblStatus.Text = "Iniciando.....";
            ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_WERKS")).Text = txtCentro.Text;
            ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_QENTST")).Text = this.date1.Text.Replace("/",".");
            ((GuiComboBox)Session.FindById("wnd[0]/usr/cmbP_BOLE")).Key = "000000000000000023";
            ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();

            string CANAPROP;
            string CANAFORN;
            string CANATOTAL;
            string ACUCAREXTRA2A;
            string VALORMESPROP;
            string VALORMESFORN;
            string VALORMESTOTAL;
            string CANAESTOQUE;
            string CANAESTOQUESEMANA;
            string VHP;
            string TOTALSAFRAVHP;
            string BONSUCROVHP;
            string BONSUCROACUCAR;
            string QUINZENA;

            CANAPROP = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(1, "VALORDIA");
            CANAFORN = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(2, "VALORDIA");
            CANATOTAL = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(3, "VALORSAFRAACT");
            VALORMESPROP = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(1, "VALORMES");
            VALORMESFORN = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(2, "VALORMES");
            VALORMESTOTAL = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(3, "VALORMES");

            ACUCAREXTRA2A = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(18, "VALORDIA");
            CANAESTOQUE = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(10, "VALORDIA");
            CANAESTOQUESEMANA = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(10, "VALORSEMANA");
            VHP = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(16, "VALORDIA");
            TOTALSAFRAVHP = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(16, "VALORSAFRAACT");
            BONSUCROACUCAR = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(20, "VALORDIA");
            BONSUCROVHP = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")).GetCellValue(19, "VALORDIA");

            if (monthCalendar1.SelectionEnd.ToString("dd") == "01" || monthCalendar1.SelectionEnd.ToString("dd") == "02" || monthCalendar1.SelectionEnd.ToString("dd") == "03" || monthCalendar1.SelectionEnd.ToString("dd") == "04" || monthCalendar1.SelectionEnd.ToString("dd") == "05" || monthCalendar1.SelectionEnd.ToString("dd") == "06" || monthCalendar1.SelectionEnd.ToString("dd") == "07" || monthCalendar1.SelectionEnd.ToString("dd") == "08" || monthCalendar1.SelectionEnd.ToString("dd") == "09" || monthCalendar1.SelectionEnd.ToString("dd") == "10" || monthCalendar1.SelectionEnd.ToString("dd") == "11" || monthCalendar1.SelectionEnd.ToString("dd") == "12" || monthCalendar1.SelectionEnd.ToString("dd") == "13" || monthCalendar1.SelectionEnd.ToString("dd") == "14" || monthCalendar1.SelectionEnd.ToString("dd") == "15")
            {
                QUINZENA = "1";
                /***************************************************************************************/
                MySqlCommand cmd = new MySqlCommand("INSERT INTO `tb_boletim` (`CanaProp`, `CanaForn`, `CanaTotal`, `ValorMesProp`, `ValorMesForn`, `ValorMesTotal`, `CanaEstoqueVLDIA`, `CanaEstoqueVLSEM`, `VHP`, `EspExtra2-A`, `TotalSafraVHP`, `BonsucroVHP`, `BonsicrpACUCAR`, `Quinzena`, `datadoDia`, `dataImport`) " +
            "VALUES " +
            "('" + CANAPROP.Trim() + "', '" + CANAFORN.Trim() + "', '" + CANATOTAL.Trim() + "', '" + VALORMESPROP.Trim() + "', '" + VALORMESFORN.Trim() + "', '" + VALORMESTOTAL.Trim() + "', '" + CANAESTOQUE.Trim('-') + "', '" + CANAESTOQUESEMANA.Trim('-') + "', '" + VHP.Trim() + "', '" + ACUCAREXTRA2A.Trim() + "', '" + TOTALSAFRAVHP.Trim() + "', '" + BONSUCROACUCAR.Trim() + "', '" + BONSUCROVHP.Trim() + "', '" + QUINZENA + "', '" + monthCalendar1.SelectionEnd.ToString("yyyy-MM-dd") + "', NOW())", ConexaoDados.GetConnectionFaturameto());
                cmd.ExecuteNonQuery();
                /****************************************************************************/
            }
            if (monthCalendar1.SelectionEnd.ToString("dd") == "16" || monthCalendar1.SelectionEnd.ToString("dd") == "17" || monthCalendar1.SelectionEnd.ToString("dd") == "18" || monthCalendar1.SelectionEnd.ToString("dd") == "19" || monthCalendar1.SelectionEnd.ToString("dd") == "20" || monthCalendar1.SelectionEnd.ToString("dd") == "21" || monthCalendar1.SelectionEnd.ToString("dd") == "22" || monthCalendar1.SelectionEnd.ToString("dd") == "23" || monthCalendar1.SelectionEnd.ToString("dd") == "24" || monthCalendar1.SelectionEnd.ToString("dd") == "25" || monthCalendar1.SelectionEnd.ToString("dd") == "26" || monthCalendar1.SelectionEnd.ToString("dd") == "27" || monthCalendar1.SelectionEnd.ToString("dd") == "28" || monthCalendar1.SelectionEnd.ToString("dd") == "29" || monthCalendar1.SelectionEnd.ToString("dd") == "30" || monthCalendar1.SelectionEnd.ToString("dd") == "31")
            {
                QUINZENA = "2";
                /***************************************************************************************/
                MySqlCommand cmd = new MySqlCommand("INSERT INTO `tb_boletim` (`CanaProp`, `CanaForn`, `CanaTotal`, `ValorMesProp`, `ValorMesForn`, `ValorMesTotal`, `CanaEstoqueVLDIA`, `CanaEstoqueVLSEM`, `VHP`, `EspExtra2-A`, `TotalSafraVHP`, `BonsucroVHP`, `BonsicrpACUCAR`, `Quinzena`, `datadoDia`, `dataImport`) " +
                    "VALUES " +
                    "('" + CANAPROP.Trim() + "', '" + CANAFORN.Trim() + "', '" + CANATOTAL.Trim() + "', '" + VALORMESPROP.Trim() + "', '" + VALORMESFORN.Trim() + "', '" + VALORMESTOTAL.Trim() + "', '" + CANAESTOQUE.Trim('-') + "', '" + CANAESTOQUESEMANA.Trim('-') + "', '" + VHP.Trim() + "', '" + ACUCAREXTRA2A.Trim() + "', '" + TOTALSAFRAVHP.Trim() + "', '" + BONSUCROACUCAR.Trim() + "', '" + BONSUCROVHP.Trim() + "', '" + QUINZENA + "', '" + monthCalendar1.SelectionEnd.ToString("yyyy-MM-dd") + "', NOW())", ConexaoDados.GetConnectionFaturameto());
                cmd.ExecuteNonQuery();
                /****************************************************************************/
            }


            LblStatus.Text = "Concluido.....";
            Session.SendCommand("/N");
            guiWindow.SendVKey(0);

            //MessageBox.Show(ErroProg.Message);
            //LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
            //break;


        }
        private void Boletim()
        {
            DeleteData();
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "dd.MM.yyyy";
            table.Rows.Clear();
            BaixarBoletim();
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "dd/MM/yyyy";
        }
        private void RdBoletim_CheckedChanged(object sender, EventArgs e)
        {
            monthCalendar1.MaxSelectionCount = 1;
        }
        private void RDCompleto_CheckedChanged(object sender, EventArgs e)
        {
            monthCalendar1.MaxSelectionCount = 7;
        }
        private void RDSeparado_CheckedChanged(object sender, EventArgs e)
        {
        }
        private void RDPosicAcucar_CheckedChanged(object sender, EventArgs e)
        {
            monthCalendar1.MaxSelectionCount = 31;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OvAlter = 1;
            OVSeparar();
        }
        private void cmbSafra_DockChanged(object sender, EventArgs e)
        {
        }
        private void cmbSafra_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox1.Text = cmbSafra.ValueMember;
            }
        }

        /*private void VerificarOrdem()
        {
            string OrdemVenda;
            string TipodeOrdem;
            //col_TipoOV
            MySqlCommand com = new MySqlCommand();
            com.Connection = ConexaoDados.GetConnectionFaturameto();
            com.CommandText = "SELECT * FROM tb_ordem_venda WHERE Data_doc BETWEEN '" + monthCalendar1.SelectionStart.ToString("yyyy/MM/dd").Replace("/", ".") + "' AND '" + monthCalendar1.SelectionEnd.ToString("yyyy/MM/dd").Replace("/", ".") + "' ";
            MySqlDataReader dr = com.ExecuteReader();
            while (dr.Read())
            {
                OrdemVenda = dr["Doc_SD"].ToString();
                TipodeOrdem = dr["col_TipoOV"].ToString();
                MessageBox.Show(OrdemVenda +" "+ TipodeOrdem);
                //break;
            }
        }*/
    }
}
