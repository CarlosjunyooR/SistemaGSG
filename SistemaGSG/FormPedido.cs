using MetroFramework;
using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Data;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormPedido : MetroFramework.Forms.MetroForm
    {

        string usuarioLogado = System.Environment.UserName;
        public bool IsPostBack { get; private set; }
        public string vbCr { get; private set; }
        public FormPedido()
        {
            InitializeComponent();
        }
        string MesSelecionado;

        private void groupBox7_Enter(object sender, EventArgs e)
        {
        }
        private void metroCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (chboxMigo.Checked)
            {
                //Abre Campos
                dtDoc.Enabled = true;
                dtLanc.Enabled = true;
                txtNf.Enabled = true;
                txtPedido.Enabled = true;
                //Fecha Campos
                btnPedido.Enabled = false;
            }
            else
            {
                //Fecha Campos
                dtDoc.Enabled = false;
                dtLanc.Enabled = false;
                txtNf.Enabled = false;
                txtPedido.Enabled = false;

                //Abre Campos
                btnPedido.Enabled = true;
            }
        }
        private void FormPedido_Load(object sender, EventArgs e)
        {
            Filtrar();
            dtDoc.Enabled = false;
            dtLanc.Enabled = false;
            txtNf.Enabled = false;
            txtPedido.Enabled = false;
            dtMiroFatura.Enabled = false;
            dateTimePicker2.Enabled = false;

            if (chboxMiro.Checked)
            {
                //Fecha Campos
                dtDoc.Enabled = false;
                dtLanc.Enabled = false;
                txtNf.Enabled = false;
                txtPedido.Enabled = false;
                chboxMigo.Enabled = false;
                //Abre Campos
                dtMiroFatura.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
            else
            {
                //Fecha Campos
                dtMiroFatura.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }
        private void chboxMiro_CheckedChanged(object sender, EventArgs e)
        {
            if (chboxMiro.Checked)
            {
                //Fecha Campos
                dtDoc.Enabled = false;
                dtLanc.Enabled = false;
                txtNf.Enabled = false;
                txtPedido.Enabled = false;
                chboxMigo.Enabled = false;
                //Abre Campos
                dtMiroFatura.Enabled = true;
                dateTimePicker2.Enabled = true;
            }
            else
            {
                //Abre Campos
                dtDoc.Enabled = true;
                dtLanc.Enabled = true;
                txtNf.Enabled = true;
                txtPedido.Enabled = true;
                chboxMigo.Enabled = true;
                //Fecha Campos
                dtMiroFatura.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }
        private void criar_pedido()
        {

            if (string.IsNullOrEmpty(MesRef.Text))
            {
                MetroMessageBox.Show(this, "Preencha o mês de referencia.", "Informação!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                int NumLinhas = dataGridView1.RowCount;
                if (NumLinhas > 0)
                {
                    if (MetroMessageBox.Show(this, "Pedido com " + NumLinhas + " itens, deseja continuar?.", "Informação!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        //Pega a tela de execução do Windows
                        CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
                        //Pega a entrada ROT para o SAP Gui para conectar-se ao COM|SAPGUI
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
                        Session.SendCommand("/NME21N");
                        //Inicia a Barra de Progresso em 25%
                        ProgBar.Value = 25;
                        //Tecla Enter
                        string Result = MesRef.Text.Substring(0, 2);
                        label2.Text = NumLinhas.ToString();
                        guiWindow.SendVKey(0);
                        if (Result == "01")
                        {
                            MesSelecionado = "Janeiro";
                        }
                        if (Result == "02")
                        {
                            MesSelecionado = "Fevereiro";
                        }
                        if (Result == "03")
                        {
                            MesSelecionado = "Março";
                        }
                        if (Result == "04")
                        {
                            MesSelecionado = "Abril";
                        }
                        if (Result == "05")
                        {
                            MesSelecionado = "Maio";
                        }
                        if (Result == "06")
                        {
                            MesSelecionado = "Junho";
                        }
                        if (Result == "07")
                        {
                            MesSelecionado = "Julho";
                        }
                        if (Result == "08")
                        {
                            MesSelecionado = "Agosto";
                        }
                        if (Result == "09")
                        {
                            MesSelecionado = "Setembro";
                        }
                        if (Result == "10")
                        {
                            MesSelecionado = "Outubro";
                        }
                        if (Result == "11")
                        {
                            MesSelecionado = "Novembro";
                        }
                        if (Result == "12")
                        {
                            MesSelecionado = "Dezembro";
                        }
                        int Container = 1;
                        while (Container < 99)
                        {
                            try
                            {
                                Container++;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + Container + "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).Text = "2000000246";
                                break;
                            }
                            catch
                            {

                            }
                        }
                        //Cód. Fornecedor//
                        ((GuiComboBox)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + Container + "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART")).Key = "NB";
                        //Tecla Enter
                        guiWindow.SendVKey(0);
                        // Modifica o tipo de formato na data referênte ao Banco de Dados Ex.: 2020/06/27 para 27/06/2020.
                        dtDoc.Text = dataGridView1.Rows[0].Cells[18].Value.ToString();
                        dtDoc.Format = DateTimePickerFormat.Custom;
                        dtDoc.CustomFormat = "dd/MM/yyyy";

                        //Dados Organizacionais
                        int OrgCont = 1;
                        while (OrgCont < 99)
                        {
                            try
                            {
                                OrgCont++;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG")).Text = "1000";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP")).Text = "400";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS")).Text = "USGA";
                                break;
                            }
                            catch
                            {

                            }
                        }
                        //Seleciona a aba texto e adiciona a nota fiscal e data
                        int TextPedido = 1;
                        while (TextPedido < 99)
                        {
                            try
                            {
                                TextPedido++;
                                ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + TextPedido + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3")).Select();
                                break;
                            }
                            catch
                            {

                            }
                        }
                        int TextPedido2 = 1;
                        while (TextPedido < 99)
                        {
                            try
                            {
                                TextPedido2++;
                                ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + TextPedido2 + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")).Text = "REF. NOTAS FISCAIS AGRUPADAS DO MÊS DE " + MesSelecionado.ToUpper() + "." + vbCr + "";
                                break;
                            }
                            catch
                            {

                            }
                        }

                        //Conta quantas linhas(itens) tem na nota fiscal referida
                        int countg = dataGridView1.RowCount;
                        //Condição para criar o pedido com 1 item e por diante
                        if (countg > 0)
                        {
                            try
                            {
                                //Barra de Progresso
                                ProgBar.Value = 35;
                                //Tecla Enter
                                guiWindow.SendVKey(0);
                                int ItensPedidoCont = 1;
                                while (ItensPedidoCont < 99)
                                {
                                    try
                                    {
                                        ItensPedidoCont++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dataGridView1.Rows[0].Cells[1].Value.ToString();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]")).Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = "1";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 28;
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerR = 1;
                                while (ContainerR < 99)
                                {
                                    try
                                    {
                                        ContainerR++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerR + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dataGridView1.Rows[0].Cells[5].Value.ToString();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerRImp = 1;
                                while (ContainerRImp < 99)
                                {
                                    try
                                    {
                                        ContainerRImp++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRImp + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dataGridView1.Rows[0].Cells[6].Value.ToString();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerRTab = 1;
                                while (ContainerRTab < 99)
                                {
                                    try
                                    {
                                        ContainerRTab++;
                                        ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRTab + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();

                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                int ContainerRCond = 1;
                                while (ContainerRCond < 99)
                                {
                                    try
                                    {
                                        ContainerRCond++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dataGridView1.Rows[0].Cells[7].Value.ToString();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                int SelecTextoItem = 1;
                                while (SelecTextoItem < 99)
                                {
                                    try
                                    {
                                        SelecTextoItem++;
                                        ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + SelecTextoItem + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15")).Select();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                int ItemdoPedidoTexto = 1;
                                dtDoc.Text = dataGridView1.Rows[0].Cells["emissao"].Value.ToString();
                                while (ItemdoPedidoTexto < 99)
                                {
                                    try
                                    {
                                        ItemdoPedidoTexto++;
                                        ((GuiTree)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTexto + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")).SelectedNode = "F01";
                                        ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTexto + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")).Text = "CÓDIGO ÚNICO: " + dataGridView1.Rows[0].Cells["cod_unico"].Value.ToString() + " & NOTA FISCAL Nº:" + dataGridView1.Rows[0].Cells["nfe"].Value.ToString() + " DE " + this.dtDoc.Text + " REFERENTE AO MÊS DE " + MesSelecionado.ToUpper() + "/2021.";
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                //Segundo
                                int numero = 1;
                                int item = 1;
                                int BarraProgresso = 45;
                                int ItemSequ = 2;
                                while (numero < countg)
                                {
                                    try
                                    {
                                        ProgBar.Value = BarraProgresso;

                                        int ContainerRPressButton = 1;
                                        while (ContainerRPressButton < 99)
                                        {
                                            try
                                            {
                                                ContainerRPressButton++;
                                                ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRPressButton + "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        label9.Text = ItemSequ.ToString();
                                        int ContainerRSegItnes = 1;
                                        while (ContainerRSegItnes < 99)
                                        {
                                            try
                                            {
                                                ContainerRSegItnes++;
                                                ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = item;
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }

                                        int ContSItens = 1;
                                        while (ContSItens < 99)
                                        {
                                            ContSItens++;
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dataGridView1.Rows[item].Cells[1].Value.ToString();
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,1]")).Text = dataGridView1.Rows[item].Cells[2].Value.ToString();
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = "1";
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                            break;
                                        }
                                        guiWindow.SendVKey(0);
                                        guiWindow.SendVKey(0);
                                        int ContainerRSegD = 1;
                                        while (ContainerRSegD < 99)
                                        {
                                            try
                                            {
                                                ContainerRSegD++;
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegD + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dataGridView1.Rows[item].Cells[5].Value.ToString();
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        guiWindow.SendVKey(0);

                                        int ContainerRSegDi = 1;
                                        while (ContainerRSegDi < 99)
                                        {
                                            try
                                            {
                                                ContainerRSegDi++;
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dataGridView1.Rows[item].Cells[6].Value.ToString();
                                                ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        int ContainerRSegDii = 1;
                                        while (ContainerRSegDii < 99)
                                        {
                                            try
                                            {
                                                ContainerRSegDii++;
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dataGridView1.Rows[item].Cells[7].Value.ToString();
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        int SelecTextoItemN = 1;
                                        while (SelecTextoItemN < 99)
                                        {
                                            try
                                            {
                                                SelecTextoItemN++;
                                                ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + SelecTextoItemN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15")).Select();
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        int ItemdoPedidoTextoN = 1;
                                        dtDoc.Text = dataGridView1.Rows[item].Cells["emissao"].Value.ToString();
                                        while (ItemdoPedidoTextoN < 99)
                                        {
                                            try
                                            {
                                                ItemdoPedidoTextoN++;
                                                ((GuiTree)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")).SelectedNode = "F01";
                                                ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")).Text = "CÓDIGO ÚNICO: " + dataGridView1.Rows[item].Cells["cod_unico"].Value.ToString() + " & NOTA FISCAL Nº:" + dataGridView1.Rows[item].Cells["nfe"].Value.ToString() + " DE " + this.dtDoc.Text + " REFERENTE AO MÊS DE " + MesSelecionado.ToUpper() + "/2021.";
                                                break;
                                            }
                                            catch
                                            {

                                            }
                                        }
                                        guiWindow.SendVKey(0);
                                        item++;
                                        numero++;
                                        BarraProgresso++;
                                        ItemSequ++;
                                    }
                                    catch
                                    {
                                        break;
                                    }
                                }
                                //Barra de Progresso
                                ProgBar.Value = 929;
                                if (MessageBox.Show("Valores Corretos?!", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    ///Pressiona o Botão Gravar
                                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                                    //Pega a Barra de Status do SAP
                                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                                    //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                                    try
                                    {

                                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET pedido='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE Mes_ref='" + MesRef.Text + "' AND err_col ='1'", ConexaoDados.GetConnectionEquatorial());
                                        cmd.ExecuteNonQuery();
                                        ConexaoDados.GetConnectionEquatorial().Close();
                                    }
                                    catch (MySqlException ErroMysql)
                                    {
                                        MessageBox.Show(ErroMysql.Message);
                                    }
                                }
                            }
                            catch (Exception Erro)
                            {
                                MessageBox.Show(Erro.Message);
                            }
                        }
                        //Finaliza com 100% a Barra de Progresso
                        ProgBar.Value = 1000;
                    }
                }
                else
                {
                    int NumLinhasDt2 = dataGridView2.RowCount;
                    if (MetroMessageBox.Show(this, "Pedido com " + NumLinhasDt2 * 2 + " itens, deseja continuar?.", "Informação!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {

                        //Pega a tela de execução do Windows
                        CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
                        //Pega a entrada ROT para o SAP Gui para conectar-se ao COM|SAPGUI
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
                        Session.SendCommand("/NME21N");
                        //Inicia a Barra de Progresso em 25%
                        ProgBar.Value = 25;
                        //Tecla Enter
                        string Result = MesRef.Text.Substring(0, 2);
                        label2.Text = NumLinhasDt2.ToString();
                        guiWindow.SendVKey(0);
                        if (string.IsNullOrWhiteSpace(MesRef.Text))
                        {

                        }
                        else
                        {
                            if (Result == "01")
                            {
                                MesSelecionado = "Janeiro";
                            }
                            if (Result == "02")
                            {
                                MesSelecionado = "Fevereiro";
                            }
                            if (Result == "03")
                            {
                                MesSelecionado = "Março";
                            }
                            if (Result == "04")
                            {
                                MesSelecionado = "Abril";
                            }
                            if (Result == "05")
                            {
                                MesSelecionado = "Maio";
                            }
                            if (Result == "06")
                            {
                                MesSelecionado = "Junho";
                            }
                            if (Result == "07")
                            {
                                MesSelecionado = "Julho";
                            }
                            if (Result == "08")
                            {
                                MesSelecionado = "Agosto";
                            }
                            if (Result == "09")
                            {
                                MesSelecionado = "Setembro";
                            }
                            if (Result == "10")
                            {
                                MesSelecionado = "Outubro";
                            }
                            if (Result == "11")
                            {
                                MesSelecionado = "Novembro";
                            }
                            if (Result == "12")
                            {
                                MesSelecionado = "Dezembro";
                            }
                        }
                        int Container = 1;
                        while (Container < 99)
                        {
                            try
                            {
                                Container++;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + Container + "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).Text = "2000000246";
                                break;
                            }
                            catch
                            {

                            }
                        }
                        //Cód. Fornecedor//
                        ((GuiComboBox)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + Container + "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART")).Key = "NB";
                        //Tecla Enter
                        guiWindow.SendVKey(0);
                        // Modifica o tipo de formato na data referênte ao Banco de Dados Ex.: 2020/06/27 para 27/06/2020.
                        dtDoc.Text = dataGridView2.Rows[0].Cells["emissao"].Value.ToString();
                        dtDoc.Format = DateTimePickerFormat.Custom;
                        dtDoc.CustomFormat = "dd/MM/yyyy";

                        //Dados Organizacionais
                        int OrgCont = 1;
                        while (OrgCont < 99)
                        {
                            try
                            {
                                OrgCont++;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG")).Text = "1000";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP")).Text = "400";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + OrgCont + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS")).Text = "USGA";
                                break;
                            }
                            catch
                            {

                            }
                        }
                        //Seleciona a aba texto e adiciona a nota fiscal e data
                        int TextPedido = 1;
                        while (TextPedido < 99)
                        {
                            try
                            {
                                TextPedido++;
                                ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + TextPedido + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3")).Select();
                                break;
                            }
                            catch
                            {

                            }
                        }
                        int TextPedido2 = 1;
                        while (TextPedido < 99)
                        {
                            try
                            {
                                TextPedido2++;
                                ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + TextPedido2 + "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")).Text = "REF. NOTAS FISCAIS AGRUPADAS DO MÊS DE " + MesSelecionado.ToUpper() + "/" + MesRef.Text.Substring(3, 4) + "." + vbCr + "";
                                break;
                            }
                            catch
                            {

                            }
                        }

                        //Conta quantas linhas(itens) tem na nota fiscal referida
                        int countg = dataGridView2.RowCount;
                        //Condição para criar o pedido com 1 item e por diante
                        if (countg > 0)
                        {
                            try
                            {
                                //Barra de Progresso
                                ProgBar.Value = 35;
                                //Tecla Enter
                                guiWindow.SendVKey(0);
                                int ItensPedidoCont = 1;
                                while (ItensPedidoCont < 99)
                                {
                                    try
                                    {
                                        ItensPedidoCont++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dataGridView2.Rows[0].Cells[1].Value.ToString();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]")).Text = dataGridView2.Rows[0].Cells[2].Value.ToString();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = "1";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItensPedidoCont + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 28;
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerR = 1;
                                while (ContainerR < 99)
                                {
                                    try
                                    {
                                        ContainerR++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerR + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dataGridView2.Rows[0].Cells[5].Value.ToString();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerRImp = 1;
                                while (ContainerRImp < 99)
                                {
                                    try
                                    {
                                        ContainerRImp++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRImp + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dataGridView2.Rows[0].Cells[6].Value.ToString();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                guiWindow.SendVKey(0);
                                int ContainerRTab = 1;
                                while (ContainerRTab < 99)
                                {
                                    try
                                    {
                                        ContainerRTab++;
                                        ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRTab + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();

                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                int ContainerRCond = 1;
                                while (ContainerRCond < 99)
                                {
                                    try
                                    {
                                        ContainerRCond++;
                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dataGridView2.Rows[0].Cells[7].Value.ToString();
                                        ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 9;
                                        string TipoIva = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,3]")).Text;
                                        if (TipoIva.Trim() == "ZCOI")
                                        {
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,5]")).Text = "ZICC";
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]")).Text = dataGridView2.Rows[0].Cells["vl_icms"].Value.ToString();
                                        }
                                        else
                                        {
                                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,3]")).Text = dataGridView2.Rows[0].Cells["vl_icms"].Value.ToString();
                                        }
                                        ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRCond + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 0;

                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                guiWindow.SendVKey(0);
                                int SelecTextoItem = 1;
                                while (SelecTextoItem < 99)
                                {
                                    try
                                    {
                                        SelecTextoItem++;
                                        ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + SelecTextoItem + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15")).Select();
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                int ItemdoPedidoTexto = 1;
                                dtDoc.Text = dataGridView2.Rows[0].Cells["emissao"].Value.ToString();
                                while (ItemdoPedidoTexto < 99)
                                {
                                    try
                                    {
                                        ItemdoPedidoTexto++;
                                        ((GuiTree)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTexto + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")).SelectedNode = "F01";
                                        ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTexto + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")).Text = "CÓDIGO ÚNICO: " + dataGridView2.Rows[0].Cells["cod_unico"].Value.ToString() + " & NOTA FISCAL Nº:" + dataGridView2.Rows[0].Cells["nfe"].Value.ToString() + " DE " + this.dtDoc.Text + " REFERENTE AO MÊS DE " + MesSelecionado.ToUpper() + "/2021.";
                                        break;
                                    }
                                    catch
                                    {

                                    }
                                }
                                //Segundo Item
                                int numero = 1;
                                int item = 0;
                                int BarraProgresso = 45;
                                int ItemSequ = 2;
                                int ItemSelecionado = 2;
                                int VerticalScrollbar = 1;
                                while (numero < countg)
                                {
                                    if (ItemSelecionado == 1)
                                    {
                                        try
                                        {
                                            ProgBar.Value = BarraProgresso;

                                            int ContainerRPressButton = 1;
                                            while (ContainerRPressButton < 99)
                                            {
                                                try
                                                {
                                                    ContainerRPressButton++;
                                                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRPressButton + "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            label9.Text = ItemSequ.ToString();
                                            int ContainerRSegItnes = 1;
                                            while (ContainerRSegItnes < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegItnes++;
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = VerticalScrollbar;
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }

                                            int ContSItens = 1;
                                            while (ContSItens < 99)
                                            {
                                                ContSItens++;
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dataGridView2.Rows[item].Cells[1].Value.ToString();
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,1]")).Text = dataGridView2.Rows[item].Cells[2].Value.ToString();
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = "1";
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                                break;
                                            }
                                            guiWindow.SendVKey(0);
                                            guiWindow.SendVKey(0);
                                            int ContainerRSegD = 1;
                                            while (ContainerRSegD < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegD++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegD + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dataGridView2.Rows[item].Cells[5].Value.ToString();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            guiWindow.SendVKey(0);

                                            int ContainerRSegDi = 1;
                                            while (ContainerRSegDi < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegDi++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dataGridView2.Rows[item].Cells[6].Value.ToString();
                                                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int ContainerRSegDii = 1;
                                            while (ContainerRSegDii < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegDii++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dataGridView2.Rows[item].Cells["base_calculo"].Value.ToString();
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 9;
                                                    string TipoIva = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,3]")).Text;
                                                    if (TipoIva.Trim() == "ZCOI")
                                                    {
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,5]")).Text = "ZICC";
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]")).Text = dataGridView2.Rows[item].Cells["vl_icms"].Value.ToString();
                                                    }
                                                    else
                                                    {
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,3]")).Text = dataGridView2.Rows[item].Cells["vl_icms"].Value.ToString();
                                                    }
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 0;

                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int SelecTextoItemN = 1;
                                            while (SelecTextoItemN < 99)
                                            {
                                                try
                                                {
                                                    SelecTextoItemN++;
                                                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + SelecTextoItemN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15")).Select();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int ItemdoPedidoTextoN = 1;
                                            dtDoc.Text = dataGridView2.Rows[item].Cells["emissao"].Value.ToString();
                                            while (ItemdoPedidoTextoN < 99)
                                            {
                                                try
                                                {
                                                    ItemdoPedidoTextoN++;
                                                    ((GuiTree)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")).SelectedNode = "F01";
                                                    ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")).Text = "CÓDIGO ÚNICO: " + dataGridView2.Rows[item].Cells["cod_unico"].Value.ToString() + " & NOTA FISCAL Nº:" + dataGridView2.Rows[item].Cells["nfe"].Value.ToString() + " DE " + this.dtDoc.Text + " REFERENTE AO MÊS DE " + MesSelecionado.ToUpper() + "/" + MesRef.Text.Substring(3, 4) + ".";
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                        }
                                        catch
                                        {
                                            break;
                                        }
                                        /**********************************FIM*********************************************************/
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        ItemSelecionado = 2;///////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /***********************************************************************************************/
                                        VerticalScrollbar++;
                                    }
                                    if (ItemSelecionado == 2)
                                    {
                                        try
                                        {
                                            ProgBar.Value = BarraProgresso;

                                            int ContainerRPressButton = 1;
                                            while (ContainerRPressButton < 99)
                                            {
                                                try
                                                {
                                                    ContainerRPressButton++;
                                                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRPressButton + "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            label9.Text = ItemSequ.ToString();
                                            int ContainerRSegItnes = 1;
                                            while (ContainerRSegItnes < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegItnes++;
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = VerticalScrollbar;
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }

                                            int ContSItens = 1;
                                            while (ContSItens < 99)
                                            {
                                                ContSItens++;
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dataGridView2.Rows[item].Cells["material_dif"].Value.ToString();
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,1]")).Text = dataGridView2.Rows[item].Cells["desc_item_dif"].Value.ToString();
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = "1";
                                                ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegItnes + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                                break;
                                            }
                                            guiWindow.SendVKey(0);
                                            guiWindow.SendVKey(0);
                                            int ContainerRSegD = 1;
                                            while (ContainerRSegD < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegD++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegD + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dataGridView2.Rows[item].Cells["custo_dif"].Value.ToString();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            guiWindow.SendVKey(0);

                                            int ContainerRSegDi = 1;
                                            while (ContainerRSegDi < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegDi++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dataGridView2.Rows[item].Cells["cod_imp_dif"].Value.ToString();
                                                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDi + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int ContainerRSegDii = 1;
                                            while (ContainerRSegDii < 99)
                                            {
                                                try
                                                {
                                                    ContainerRSegDii++;
                                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dataGridView2.Rows[item].Cells["vl_dif"].Value.ToString();
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 9;
                                                    string TipoIva = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,3]")).Text;
                                                    if (TipoIva.Trim() == "ZCOI")
                                                    {
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,5]")).Text = "ZICC";
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]")).Text = dataGridView2.Rows[item].Cells["iva_dif"].Value.ToString();
                                                    }
                                                    else
                                                    {
                                                        ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,3]")).Text = dataGridView2.Rows[item].Cells["iva_dif"].Value.ToString();
                                                    }
                                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ContainerRSegDii + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN")).VerticalScrollbar.Position = 0;
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int SelecTextoItemN = 1;
                                            while (SelecTextoItemN < 99)
                                            {
                                                try
                                                {
                                                    SelecTextoItemN++;
                                                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + SelecTextoItemN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15")).Select();
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            int ItemdoPedidoTextoN = 1;
                                            dtDoc.Text = dataGridView2.Rows[item].Cells["emissao"].Value.ToString();
                                            while (ItemdoPedidoTextoN < 99)
                                            {
                                                try
                                                {
                                                    ItemdoPedidoTextoN++;
                                                    ((GuiTree)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/cntlTEXT_TYPES_0200/shell")).SelectedNode = "F01";
                                                    ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + ItemdoPedidoTextoN + "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell")).Text = "CÓDIGO ÚNICO: " + dataGridView2.Rows[item].Cells["cod_unico"].Value.ToString() + " & NOTA FISCAL Nº:" + dataGridView2.Rows[item].Cells["nfe"].Value.ToString() + " DE " + this.dtDoc.Text + " REFERENTE AO MÊS DE " + MesSelecionado.ToUpper() + "/" + MesRef.Text.Substring(3, 4) + ".";
                                                    break;
                                                }
                                                catch
                                                {

                                                }
                                            }
                                            guiWindow.SendVKey(0);
                                            VerticalScrollbar++;
                                            item++;
                                            numero++;
                                            BarraProgresso++;
                                            ItemSequ++;
                                        }
                                        catch
                                        {
                                            break;
                                        }
                                        /**********************************FIM*********************************************************/
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        ItemSelecionado = 1;///////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /////////////////////////////////////////////////////////////////////////////////////
                                        /***********************************************************************************************/
                                    }
                                }
                                //Barra de Progresso
                                ProgBar.Value = 929;
                                if (MessageBox.Show("Valores Corretos?!", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    ///Pressiona o Botão Gravar
                                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                                    //Pega a Barra de Status do SAP
                                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                                    //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                                    string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                                    try
                                    {

                                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET pedido='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE Mes_ref='" + MesRef.Text + "' AND err_col ='1'", ConexaoDados.GetConnectionEquatorial());
                                        cmd.ExecuteNonQuery();
                                        ConexaoDados.GetConnectionEquatorial().Close();
                                    }
                                    catch (MySqlException ErroMysql)
                                    {
                                        MessageBox.Show(ErroMysql.Message);
                                    }
                                }
                            }
                            catch (Exception Erro)
                            {
                                MessageBox.Show(Erro.Message);
                            }
                        }
                        //Finaliza com 100% a Barra de Progresso
                        ProgBar.Value = 1000;
                    }
                }
            }
        }
        private void criar_migo()
        {
            //Get the Windows Running Object Table
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            //Get the ROT Entry for the SAP Gui to connect to the COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Get the reference to the Scripting Engine
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Get the reference to the running SAP Application Window
            GuiApplication GuiApp = (GuiApplication)engine;
            //Get the reference to the first open connection
            GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //get the first available session
            GuiSession Session = (GuiSession)connection.Children.ElementAt(0);
            //Get the reference to the main "Frame" in which to send virtual key commands
            GuiFrameWindow guiWindow = Session.ActiveWindow;

            //Maximisa Janela
            guiWindow.Maximize();

            int countg = dataGridView1.RowCount;
            int numero = 0;
            while (numero < countg)
            {
                try
                {
                    //Abre Transação
                    Session.SendCommand("/NMIGO");

                    txtPedido.Text = dataGridView1.Rows[numero].Cells[26].Value.ToString();
                    this.dtDoc.Text = dataGridView1.Rows[numero].Cells[18].Value.ToString();
                    txtNf.Text = dataGridView1.Rows[numero].Cells[19].Value.ToString();

                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER")).Text = txtPedido.Text.Replace(" ", "");
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BLDAT")).Text = this.dtDoc.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BUDAT")).Text = this.dtLanc.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).Text = txtNf.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).CaretPosition = 8;
                    guiWindow.SendVKey(0);
                    //((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,0]")).Selected = 1;
                    //((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,1]")).Selected = 1;
                    ((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,1]")).SetFocus();
                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));

                    MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET migo='" + resultado.Split('r')[0] + "' WHERE Mes_ref='" + MesRef.Text + "' AND err_col ='1'", ConexaoDados.GetConnectionEquatorial());
                    cmd.ExecuteNonQuery();
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NOT NULL AND migo=''", ConexaoDados.GetConnectionEquatorial());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    dataGridView1.DataSource = SS;
                    ConexaoDados.GetConnectionEquatorial().Close();
                    numero++;
                }
                catch
                {
                    break;
                }
            }
        }
        private void btnMigo_Click(object sender, EventArgs e)
        {
            try
            {
                criar_migo();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Por favor Contate o Administrador do SistemaGSG!.\n'" + ex.Message + "'", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                MessageBox.Show("Fim!");
            }
        }

        private void btnFilterMigo_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NOT NULL AND migo='' AND err_col='1'", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch
            {
                MessageBox.Show("Não Existe Itens sem Migo no Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        private void btnPedido_Click_1(object sender, EventArgs e)
        {
            try
            {
                criar_pedido();
                //VerificarSAPGUI();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }
        private void btnPedidoNormal_Click(object sender, EventArgs e)
        {
            Filtrar();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NOT NULL AND miro=''", ConexaoDados.GetConnectionEquatorial());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
            }
            catch
            {
                MessageBox.Show("Não Existe Itens para Criar Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        private void btnCriarMiro_Click(object sender, EventArgs e)
        {
            //Get the Windows Running Object Table
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            //Get the ROT Entry for the SAP Gui to connect to the COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Get the reference to the Scripting Engine
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Get the reference to the running SAP Application Window
            GuiApplication GuiApp = (GuiApplication)engine;
            //Get the reference to the first open connection
            GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //get the first available session
            GuiSession Session = (GuiSession)connection.Children.ElementAt(0);
            //Get the reference to the main "Frame" in which to send virtual key commands
            GuiFrameWindow guiWindow = Session.ActiveWindow;

            //Maximisa Janela
            guiWindow.Maximize();


            int countg = dataGridView1.RowCount;
            int numero = 0;
            while (numero < countg)
            {
                try
                {
                    //Abre Transação
                    Session.SendCommand("/NMIRO");
                    guiWindow.SendVKey(0);

                    this.dtMiroFatura.Text = dataGridView1.Rows[numero].Cells[18].Value.ToString();
                    //txtNfeMiro.Text = dataGridView1.Rows[numero].Cells[19].Value.ToString();
                    //txtVlMiro.Text = dataGridView1.Rows[numero].Cells[30].Value.ToString();
                    //this.dtVencimentoMiro.Text = dataGridView1.Rows[numero].Cells[23].Value.ToString();
                    txtPedido.Text = dataGridView1.Rows[numero].Cells[26].Value.ToString();
                    //txtCodUnic.Text = dataGridView1.Rows[numero].Cells[25].Value.ToString();
                    //txtMiro.Text = dataGridView1.Rows[numero].Cells[29].Value.ToString();
                    txtNf.Text = dataGridView1.Rows[numero].Cells[2].Value.ToString();

                    try
                    {
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT")).Text = this.dtMiroFatura.Text;
                        guiWindow.SendVKey(0);
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR")).Text = txtNfeMiro.Text;
                        guiWindow.SendVKey(0);
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR")).Text = txtVlMiro.Text.Replace(".", ",");
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT")).Text = txtNf.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).Text = txtPedido.Text.Replace(" ", "");
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).CaretPosition = 10;
                        guiWindow.SendVKey(0);
                        //((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT")).Text = this.dtVencimentoMiro.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZLSCH")).Text = formPag.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HBKID")).Text = bancoEmpresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HKTID")).Text = bancoEmpresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).Text = txtRefPagmto.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).CaretPosition = 8;
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER")).Text = empresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).Text = txtCodUnic.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE")).Text = "S1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).CaretPosition = 9;
                        guiWindow.SendVKey(0);
                        ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[21]")).Press();
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).Text = txtMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).CaretPosition = 16;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(3);

                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                        GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                        string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));
                        //MessageBox.Show(resultado.Replace(" ", "").Replace("f", "").Replace("o", ""));

                        //MySqlCommand cmd = new MySqlCommand("UPDATE ``tb_boleto`` SET `miro`='"+ resultado.Replace(" fo","") +"' WHERE nfe='" + txtNfeMiro.Text + "'", ConexaoDados.GetConnectionEquatorial());
                        //cmd.ExecuteNonQuery();
                        MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NOT NULL AND miro=''", ConexaoDados.GetConnectionEquatorial());
                        DataTable SS = new DataTable();
                        ADAP.Fill(SS);
                        dataGridView1.DataSource = SS;
                        ConexaoDados.GetConnectionEquatorial().Close();
                        numero++;
                    }
                    catch (Exception)
                    {
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT")).Text = this.dtVencimentoMiro.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZLSCH")).Text = formPag.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HBKID")).Text = bancoEmpresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HKTID")).Text = bancoEmpresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).Text = txtRefPagmto.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).CaretPosition = 8;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER")).Text = empresa.Text;
                        //((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).Text = txtCodUnic.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE")).Text = "S1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).CaretPosition = 9;
                        guiWindow.SendVKey(0);
                        ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[21]")).Press();
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4")).Select();
                        //((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).Text = txtMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).CaretPosition = 16;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(3);

                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                        GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                        string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));
                        //MessageBox.Show(resultado.Replace(" ", "").Replace("f", "").Replace("o", ""));

                        //MySqlCommand cmd = new MySqlCommand("UPDATE ``tb_boleto`` SET `miro`='"+ resultado.Replace(" ","").Replace("f","").Replace("o","") +"' WHERE nfe='" + txtNfeMiro.Text + "'", ConexaoDados.GetConnectionEquatorial());
                        //cmd.ExecuteNonQuery();
                        MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NOT NULL AND miro=''", ConexaoDados.GetConnectionEquatorial());
                        DataTable SS = new DataTable();
                        ADAP.Fill(SS);
                        dataGridView1.DataSource = SS;
                        ConexaoDados.GetConnectionEquatorial().Close();
                        numero++;
                    }
                }
                catch (Exception Err)
                {
                    MessageBox.Show(Err.Message);
                    break;
                }
            }
        }
        private void Filtrar()
        {
            try
            {
                if (string.IsNullOrEmpty(MesRef.Text))
                {
                    MetroMessageBox.Show(this, "Preencha o mês de referencia.", "Informação!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {

                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NULL AND err_col = '1' AND Mes_ref='" + MesRef.Text + "'", ConexaoDados.GetConnectionEquatorial());
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    dataGridView1.DataSource = SS;

                    MySqlDataAdapter ADAP2 = new MySqlDataAdapter("SELECT * FROM `tb_boleto` WHERE pedido IS NULL AND err_col = '2' AND Mes_ref='" + MesRef.Text + "'", ConexaoDados.GetConnectionEquatorial());
                    DataTable SS2 = new DataTable();
                    ADAP2.Fill(SS2);
                    dataGridView2.DataSource = SS2;
                    ConexaoDados.GetConnectionEquatorial().Close();
                }
            }
            catch
            {
                MetroMessageBox.Show(this, "Não Existe Itens para Criar Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        private void btnVoltar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main frm_Main = new frm_Main();
                frm_Main.Show();
                this.Visible = false;
            }
        }
        public static GuiApplication SapGuiApp { get; set; }
        public static GuiConnection SapConnection { get; set; }
        public static GuiSession SapSession { get; set; }
        public static void OpenSAP()
        {
            GuiApplication Application;
            GuiConnection Connection;
            GuiSession Session;
            Application = (GuiApplication)Activator.CreateInstance(Type.GetTypeFromProgID("SapGui.ScriptingCtrl.1"));
            // How do I find the connection string that I use to connect to SAP?
            Connection = Application.OpenConnectionByConnectionString("/H/200.143.105.131/S/3200", false, true);
            Session = (GuiSession)Connection.Sessions.Item(0);
            Session.TestToolMode = 1;
            ((GuiTextField)Session.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField")).Text = "400";
            ((GuiTextField)Session.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField")).Text = "bjunior";
            ((GuiTextField)Session.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField")).Text = "02984646#Lua";
            ((GuiTextField)Session.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField")).Text = "PT";
            // Press the green checkmark button which is about the same as the enter key 
            GuiButton btn = (GuiButton)Session.ActiveWindow.FindByName("btn[0]", "GuiButton");
            btn.SetFocus();
            btn.Press();
        }
        private void btnSAP_Click(object sender, EventArgs e)
        {
            OpenSAP();
        }
        private void VerificarSAPGUI()
        {
            //Get the Windows Running Object Table
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            //Get the ROT Entry for the SAP Gui to connect to the COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Get the reference to the Scripting Engine
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Get the reference to the running SAP Application Window
            GuiApplication GuiApp = (GuiApplication)engine;
            //Get the reference to the first open connection
            GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //get the first available session
            GuiSession Session = (GuiSession)connection.Children.ElementAt(0);
            //Get the reference to the main "Frame" in which to send virtual key commands
            GuiFrameWindow guiWindow = Session.ActiveWindow;
            //Abre Transação
            Session.SendCommand("/NME21N");
            //Inicia a Barra de Progresso em 25%
            ProgBar.Value = 25;

            int Container = 1;
            while (Container < 99)
            {
                try
                {
                    Container++;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + Container + "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).Text = "2000000246";
                    break;
                }
                catch
                {

                }
            }
            MessageBox.Show(Container.ToString());
        }
    }
}
