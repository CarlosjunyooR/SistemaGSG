using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System.IO;
using MetroFramework;

namespace SistemaGSG
{
    public partial class frmPosicaoSemana : MetroFramework.Forms.MetroForm
    {
        public frmPosicaoSemana()
        {
            InitializeComponent();
        }
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
            catch (System.Runtime.InteropServices.COMException err)
            {
                LblStatus.ForeColor = Color.Red;
                LblStatus.Text = "SAP encontra-se fechado.....";
            }
        }
        private void BaixaSAPAcucar()
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
                ((GuiTextField)Session.FindById("wnd[1]/usr/ctxtDY_FILENAME")).Text = "POSICAOACUCAR.txt";
                LblStatus.Text = "Arquivo POSICAOACUCAR.txt Baixado com sucesso!.....";
                ((GuiButton)Session.FindById("wnd[1]/tbar[0]/btn[11]")).Press();
                Session.SendCommand("/N");

                ImportarTXTAcucar();
                PreencherTextBoxAcucar();
            }
            catch (System.Runtime.InteropServices.COMException err)
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
            MySqlCommand comandDell = new MySqlCommand("DELETE FROM tb_saida_semana WHERE DATA_EMISS BETWEEN '"+this.date1.Text+"' AND '"+this.date2.Text+"'", ConexaoDados.GetConnectionFaturameto());
            comandDell.ExecuteNonQuery();
            MySqlCommand comandDelete = new MySqlCommand("DELETE FROM tb_filaphp WHERE dataPeriodo BETWEEN '" + this.date1.Text + "' AND '" + this.date2.Text + "'", ConexaoDados.GetConnectionFaturameto());
            comandDelete.ExecuteNonQuery();
        }
        private void DeleteBetweenAcucar()
        {
            date1.Format = DateTimePickerFormat.Custom;
            date1.CustomFormat = "yyyy-MM-dd";
            date2.Format = DateTimePickerFormat.Custom;
            date2.CustomFormat = "yyyy-MM-dd";
            MySqlCommand comandDell = new MySqlCommand("DELETE FROM tb_ordem_venda WHERE Data_doc BETWEEN '" + this.date1.Text + "' AND '" + this.date2.Text + "'", ConexaoDados.GetConnectionFaturameto());
            comandDell.ExecuteNonQuery();
            //MySqlCommand comandDelete = new MySqlCommand("DELETE FROM tb_filaphp WHERE dataPeriodo BETWEEN '" + this.date1.Text + "' AND '" + this.date2.Text + "'", ConexaoDados.GetConnectionFaturameto());
            //comandDelete.ExecuteNonQuery();
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

            string[] lines = File.ReadAllLines(@"C:\ArquivosSAP\POSICAOACUCAR.txt", Encoding.UTF7);
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
                    this._28.Text = DT_SAP.Rows[numero].Cells[28].Value.ToString().Replace(".","/").Trim();
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
                catch(Exception ErroProg)
                {
                    MessageBox.Show(ErroProg.Message);
                    LblStatus.Text = "Algo de errado aconteceu print a tela e envie para o administrador!.....";
                    break;
                }
            }
            ProgressBar.Value = 1000;
            LblStatus.Text = "Pronto processo finalizado.....";
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
                            MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_ordem_venda (`Doc_SD`, `Itm`, `Div_ITEM`, `Denominacao`, `TpDV`, `Data_doc`, `Qtd_conf`, `N_pedido`, `Criado_a`, `Qtd_ordem`, `Dep`, `UMB`, `Nome_1`, `Preco_liq`, `safra`, `UM`, `Val_liq`, `MATERIAL`) " +
                            "VALUES " +
                            "('" + DocOrdem + "','" + Itm + "','" + Div + "','" + Denominacao + "','" + TpDv + "','" + _28.Text + "','" + QtdConf.Replace(".","").Replace(",",".") + "','" + Pedido + "', NULL, NULL, NULL,  '" + UMB + "', NULL, NULL,NULL, NULL, NULL,'" + MATERIAL + "')", ConexaoDados.GetConnectionFaturameto());
                            cmd.ExecuteNonQuery();
                            cmd.CommandTimeout = 120; //default 30 segundos
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
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100001")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100002")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100014")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100015")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100035")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100141")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100145")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    if (_30.Text == "100180")
                    {
                        if (_16.Text == "6923/AA")
                        {

                        }
                        else
                        {
                            if (_16.Text == "5923/AA")
                            {

                            }
                            else
                            {
                                MySqlCommand cmd = new MySqlCommand("INSERT INTO tb_saida_semana (`DOC`, `ORG`, `CANAL`, `SETOR`, `USER_SAP`, `CENTRO`, `COD_EMPR`, `NOME`, `CNPJ`, `REC_MERC`, `EMISSO_MER`, `COD_RECEB`, `CNPJ_RECEB`, `CIDADE`, `ESTADO`, `CFOP`, `DESCRICAO`, `PEDIDO`, `ORDEM`, `TIPO_ORDEM`, `FATURA`, `TIPO_FAT`, `NFE_NUM`, `NF`, `SERIE`, `TIPO`, `CANCELADA`, `DATA_EMISS`, `GRUPO_MERC`, `MATERIAL`, `DESCRICAO_MAT`, `LOTE`, `UNIDADE`, `QUANTIDADE`, `VL_LIQUIDO`, `VL_BRUTO`, `COD_REP`, `REPRESENTANTE`, `TRANSPORTADORA`, `ACESSO`, `LAUDO`, `SAFRA`, `LOTE_MANUAL`, `DEPOSITO`, `TIPO_EMB`, `QTD_EMB`, `DATA_EMISS_FIM`, `col_status`) " +
                            "VALUES " +
                            "('" + _1.Text + "','" + _2.Text + "','" + _3.Text + "','" + _4.Text + "','" + _5.Text + "','" + _6.Text + "','" + _7.Text + "','" + _8.Text + "','" + _9.Text + "','" + _10.Text + "','" + _11.Text + "','" + _12.Text + "','" + _13.Text + "','" + _14.Text + "','" + _15.Text + "','" + _16.Text.Trim() + "','" + _17.Text + "','" + _18.Text + "','" + _19.Text + "','" + _20.Text + "','" + _21.Text + "','" + _22.Text + "','" + _23.Text + "','" + _24.Text + "','" + _25.Text + "','" + _26.Text + "','" + _27.Text + "','" + _28.Text + "','" + _29.Text + "','" + _30.Text + "','" + _31.Text + "','" + _32.Text + "','" + _33.Text + "','" + _34.Text.Replace(".", "").Replace(",", ".") + "','" + _35.Text.Replace(".", "").Replace(",", ".") + "','" + _46.Text.Replace(".", "").Replace(",", ".") + "','" + _47.Text + "','" + _48.Text + "','" + _49.Text + "','" + _55.Text + "','" + _56.Text + "','" + _57.Text + "','" + _58.Text + "','" + _68.Text + "','" + _69.Text + "','" + _70.Text + "','" + this.date2.Text + "','0')", ConexaoDados.GetConnectionFaturameto());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);

                //MessageBox.Show("Erro com sua conexão, Verifique se o servidor está Online...");
            }
        }
        private void LimparTEXT()
        {
            _1.Text = "";
            _2.Text = "";
            _3.Text = "";
            _4.Text = "";
            _5.Text = "";
            _6.Text = "";
            _7.Text = "";
            _8.Text = "";
            _9.Text = "";
            _10.Text = "";
            _11.Text = "";
            _12.Text = "";
            _13.Text = "";
            _14.Text = "";
            _15.Text = "";
            _16.Text = "";
            _17.Text = "";
            _18.Text = "";
            _19.Text = "";
            _20.Text = "";
            _21.Text = "";
            _22.Text = "";
            _23.Text = "";
            _24.Text = "";
            _25.Text = "";
            _26.Text = "";
            _27.Text = "";
            this._28.Text = "";
            _29.Text = "";
            _30.Text = "";
            _31.Text = "";
            _32.Text = "";
            _33.Text = "";
            _34.Text = "";
            _35.Text = "";
            _36.Text = "";
            _37.Text = "";
            _38.Text = "";
            _39.Text = "";
            _40.Text = "";
            _41.Text = "";
            _42.Text = "";
            _43.Text = "";
            _44.Text = "";
            _45.Text = "";
            _46.Text = "";
            _47.Text = "";
            _48.Text = "";
            _49.Text = "";
            _50.Text = "";
            _51.Text = "";
            _52.Text = "";
            _53.Text = "";
            _54.Text = "";
            _55.Text = "";
            _56.Text = "";
            _57.Text = "";
            _58.Text = "";
            _59.Text = "";
            _60.Text = "";
            _61.Text = "";
            _62.Text = "";
            _63.Text = "";
            _64.Text = "";
            _65.Text = "";
            _66.Text = "";
            _67.Text = "";
            _68.Text = "";
            _69.Text = "";
            _70.Text = "";
            _71.Text = "";
            _72.Text = "";
            _73.Text = "";
            _74.Text = "";
            _75.Text = "";
            _76.Text = "";
        }
        private void button1_Click(object sender, EventArgs e)
        {
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

                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd/MM/yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd/MM/yyyy";
            }
            if (RDSeparado.Checked == true)
            {
                LimparTEXT();
                ViagensConsultaSAP();
            }
            if (RdBoletim.Checked == true)
            {
                LimparTEXT();
                Boletim();
            }
            if (RDPosicAcucar.Checked)
            {
                DeleteBetweenAcucar();

                table.Rows.Clear();
                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd.MM.yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd.MM.yyyy";

                BaixaSAPAcucar();
                ImportarTXTAcucar();

                date1.Format = DateTimePickerFormat.Custom;
                date1.CustomFormat = "dd/MM/yyyy";
                date2.Format = DateTimePickerFormat.Custom;
                date2.CustomFormat = "dd/MM/yyyy";
            }
            ConexaoDados.GetConnectionFaturameto().Close();
        }
        private void frmPosicaoSemana_Load(object sender, EventArgs e)
        {
            DataGridBancoDados();
            date1.Visible = false;
            date2.Visible = false;
           UserTXT.Text = dados.usuario;
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
                MetroMessageBox.Show(this, "Processo Finalizado.", "Sucesso!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtP_QENTST")).Text = this.date1.Text;
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
    }
}
