using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Interop.SAPFEWSELib;
using Interop.SapROTWr;
using MySql.Data.MySqlClient;
using System.IO;
using System.Text.RegularExpressions;
using ikvm.lang;
using org.ietf.jgss;
using MetroFramework;

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
        public FormPedido(string conexao)
        {
            InitializeComponent();
            txtHost.Text = conexao;
        }
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
                codigo_fornecedor.Enabled = false;
                organizacao_compras.Enabled = false;
                grupo_compras.Enabled = false;
                empresa.Enabled = false;
                categoria_pedido.Enabled = false;
                material_pedido.Enabled = false;
                descricao_item_pedido.Enabled = false;
                quantidade_item_pedido.Enabled = false;
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
                codigo_fornecedor.Enabled = true;
                organizacao_compras.Enabled = true;
                grupo_compras.Enabled = true;
                empresa.Enabled = true;
                categoria_pedido.Enabled = true;
                material_pedido.Enabled = true;
                descricao_item_pedido.Enabled = true;
                quantidade_item_pedido.Enabled = true;
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
                txtNfeMiro.Enabled = false;
                txtVlMiro.Enabled = false;
                txtMiro.Enabled = false;
                textBox9.Enabled = false;
                dtVencimentoMiro.Enabled = false;
                formPag.Enabled = false;
                bancoEmpresa.Enabled = false;
                txtRefPagmto.Enabled = false;
                txtCodUnic.Enabled = false;

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
                txtNfeMiro.Enabled = true;
                txtVlMiro.Enabled = true;
                txtMiro.Enabled = true;
                textBox9.Enabled = true;
                dtVencimentoMiro.Enabled = true;
                formPag.Enabled = true;
                bancoEmpresa.Enabled = true;
                txtRefPagmto.Enabled = true;
                txtCodUnic.Enabled = true;
            }
            else
            {
                //Fecha Campos
                dtMiroFatura.Enabled = false;
                dateTimePicker2.Enabled = false;
                txtNfeMiro.Enabled = false;
                txtVlMiro.Enabled = false;
                txtMiro.Enabled = false;
                textBox9.Enabled = false;
                dtVencimentoMiro.Enabled = false;
                formPag.Enabled = false;
                bancoEmpresa.Enabled = false;
                txtRefPagmto.Enabled = false;
                txtCodUnic.Enabled = false;
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
                txtNfeMiro.Enabled = true;
                txtVlMiro.Enabled = true;
                txtMiro.Enabled = true;
                textBox9.Enabled = true;
                dtVencimentoMiro.Enabled = true;
                formPag.Enabled = true;
                bancoEmpresa.Enabled = true;
                txtRefPagmto.Enabled = true;
                txtCodUnic.Enabled = true;
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
                txtNfeMiro.Enabled = false;
                txtVlMiro.Enabled = false;
                txtMiro.Enabled = false;
                textBox9.Enabled = false;
                dtVencimentoMiro.Enabled = false;
                formPag.Enabled = false;
                bancoEmpresa.Enabled = false;
                txtRefPagmto.Enabled = false;
                txtCodUnic.Enabled = false;
            }
        }
        private void criar_pedido()
        {
            if (string.IsNullOrEmpty(MesRef.Text))
            {
                MetroMessageBox.Show(this,"Preencha o mês de referencia.","Informação!",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {

                int NumLinhas = dtGridView1.RowCount;
                if (MetroMessageBox.Show(this, "Pedido com "+ NumLinhas +" itens, deseja continuar?.", "Informação!", MessageBoxButtons.YesNo, MessageBoxIcon.Question)==DialogResult.Yes)
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
                    Session.SendCommand(txtTrans.Text);
                    //Inicia a Barra de Progresso em 25%
                    ProgBar.Value = 25;
                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    //Cód. Fornecedor//
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).Text = codigo_fornecedor.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD")).CaretPosition = 10;
                    //Tecla Enter
                    guiWindow.SendVKey(0);
                    // Modifica o tipo de formato na data referênte ao Banco de Dados Ex.: 2020/06/27 para 27/06/2020.
                    datePickerEmissao.Text = dtGridView1.Rows[0].Cells[18].Value.ToString();
                    datePickerEmissao.Format = DateTimePickerFormat.Custom;
                    datePickerEmissao.CustomFormat = "dd/MM/yyyy";

                    //Dados Organizacionais
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG")).Text = organizacao_compras.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP")).Text = grupo_compras.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS")).Text = empresa.Text;
                    //Seleciona a aba texto e adiciona a nota fiscal e data
                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3")).Select();
                    ((GuiShell)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")).Text = "REF. NOTAS FISCAIS AGRUPADAS DO MÊS DE " + MesRef.Text + "." + vbCr + "";
                    //Conta quantas linhas(itens) tem na nota fiscal referida
                    int countg = dtGridView1.RowCount;
                    //Condição para criar o pedido com 1 item e por diante
                    if (countg > 0)
                    {
                        try
                        {
                            //Barra de Progresso
                            ProgBar.Value = 35;
                            //Tecla Enter
                            guiWindow.SendVKey(0);
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]")).Text = "K";
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]")).Text = dtGridView1.Rows[0].Cells[1].Value.ToString();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]")).Text = dtGridView1.Rows[0].Cells[2].Value.ToString();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]")).Text = "1";
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).SetFocus();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]")).CaretPosition = 28;
                            guiWindow.SendVKey(0);
                            guiWindow.SendVKey(0);

                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView1.Rows[0].Cells[6].Value.ToString();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;


                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dtGridView1.Rows[0].Cells[5].Value.ToString();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).SetFocus();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                            guiWindow.SendVKey(0);
                            guiWindow.SendVKey(0);


                            ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView1.Rows[0].Cells[7].Value.ToString();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                            ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                            guiWindow.SendVKey(0);
                            //Segundo
                            int numero = 1;
                            int item = 1;
                            int BarraProgresso = 45;
                            while (numero < countg)
                            {
                                try
                                {
                                    ProgBar.Value = BarraProgresso;
                                    ((GuiButton)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON")).Press();
                                    ((GuiTableControl)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")).VerticalScrollbar.Position = item;
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")).Text = "K";
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,1]")).Text = dtGridView1.Rows[item].Cells[15].Value.ToString();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,1]")).Text = dtGridView1.Rows[item].Cells[4].Value.ToString();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,1]")).Text = "1";
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).Text = "UNIDADE INDUSTRIAL S. GRANDE";
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).SetFocus();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]")).CaretPosition = 28;
                                    guiWindow.SendVKey(0);
                                    guiWindow.SendVKey(0);
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).Text = dtGridView1.Rows[item].Cells[7].Value.ToString();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).SetFocus();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL")).CaretPosition = 10;
                                    guiWindow.SendVKey(0);
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).Text = dtGridView1.Rows[item].Cells[16].Value.ToString();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ")).CaretPosition = 2;
                                    ((GuiTab)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8")).Select();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).Text = dtGridView1.Rows[item].Cells[5].Value.ToString();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).SetFocus();
                                    ((GuiTextField)Session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]")).CaretPosition = 16;
                                    guiWindow.SendVKey(0);
                                    item++;
                                    numero++;
                                    BarraProgresso++;
                                }
                                catch
                                {
                                    break;
                                }
                            }
                            //Barra de Progresso
                            ProgBar.Value = 92;
                            ///Pressiona o Botão Gravar
                            ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                            //Pega a Barra de Status do SAP
                            GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                            //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                            string resultado = statusbar.Text.Substring(6, statusbar.Text.IndexOf('2'));
                            //try
                            //{
                            //    MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database='" + txtDataBase.Text + "';Uid='" + txtUser.Text + "';Pwd='" + txtPass.Text + "';SslMode=none;");
                            //CONEX.Open();
                            //    MySqlCommand cmd = new MySqlCommand("UPDATE `tb_nota` SET `PEDIDO`='" + resultado.Split('º')[1].Replace(" ", "") + "' WHERE DANFE='" + txtChamarNotaFiscal.Text + "-001'", CONEX);
                            //    cmd.ExecuteNonQuery();
                            //   CONEX.Close();
                            //}
                            //catch (MySqlException ErroMysql)
                            // {
                            //   MessageBox.Show(ErroMysql.Message);
                            //}
                        }
                        catch (Exception Erro)
                        {
                            MessageBox.Show(Erro.Message);
                        }
                    }
                    //Finaliza com 100% a Barra de Progresso
                    ProgBar.Value = 100;
                    //Exibe uma menssagem de conclusão
                    MessageBox.Show("Pedido Concluido!");
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

            int countg = dtGridView1.RowCount;
            int numero = 0;
            while (numero < countg)
            {
                try
                {
                    //Abre Transação
                    Session.SendCommand("/NMIGO");

                    txtPedido.Text = dtGridView1.Rows[numero].Cells[26].Value.ToString();
                    this.dtDoc.Text = dtGridView1.Rows[numero].Cells[18].Value.ToString();
                    txtNf.Text = dtGridView1.Rows[numero].Cells[19].Value.ToString();

                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER")).Text = txtPedido.Text.Replace(" ","");
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BLDAT")).Text = this.dtDoc.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BUDAT")).Text = this.dtLanc.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).Text = txtNf.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).CaretPosition = 8;
                    guiWindow.SendVKey(0);
                    ((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,0]")).Selected = 1;
                    ((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,1]")).Selected = 1;
                    ((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,1]")).SetFocus();
                    ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                    string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));
                    MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                    CONEX.Open();
                    MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET `migo`='" + resultado.Split('r')[0] + "' WHERE nfe='" + txtNf.Text + "'", CONEX);
                    cmd.ExecuteNonQuery();
                    MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND migo=''", CONEX);
                    DataTable SS = new DataTable();
                    ADAP.Fill(SS);
                    dtGridView1.DataSource = SS;
                    CONEX.Close();
                    numero++;
                }
                catch
                {
                    break;
                }
            }
        }
        private void btnMigo_PH_Click(object sender, EventArgs e)
        {
            try
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

                //Abre Transação
                Session.SendCommand("/NMIGO");

                int numeros = 0;
                while (numeros < 1000)
                {
                    try
                    {
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER")).Text = txtPedido.Text = dtGridView1.Rows[numeros].Cells[9].Value.ToString();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BLDAT")).Text = this.dtDoc.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BUDAT")).Text = this.dtLanc.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).Text = txtNf.Text = dtGridView1.Rows[numeros].Cells[13].Value.ToString();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR")).CaretPosition = 8;
                        guiWindow.SendVKey(0);
                        //((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,0]")).Selected = true;
                        ((GuiCheckBox)Session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[2,0]")).SetFocus();
                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                        GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                        string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));

                        MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE contas_1 SET migo='" + resultado.Split('r')[0] + "' WHERE nf='" + txtNf.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        CONEX.Close();
                        numeros++;
                    }
                    catch
                    {
                        break;
                    }
                }
            }
            catch
            {

            }
            finally
            {

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                criar_pedido();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            finally
            {

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
        private void btnFilterPH_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND miro='' AND material_dif IS NOT NULL", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView1.DataSource = SS;

                int countg = dtGridView1.RowCount;
                countg--;
                item_pedido.Text = countg.ToString();
            }
            catch
            {
                MessageBox.Show("Não Existe Itens para Criar Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
        private void btnFilterMigo_Click(object sender, EventArgs e)
        {
            try
            {
                MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND migo=''", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView1.DataSource = SS;
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
                MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND miro=''", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView1.DataSource = SS;

                int countg = dtGridView1.RowCount;
                countg--;
                item_pedido.Text = countg.ToString();
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


            int countg = dtGridView1.RowCount;
            int numero = 0;
            while (numero < countg)
            {
                try
                {
                    //Abre Transação
                    Session.SendCommand("/NMIRO");
                    guiWindow.SendVKey(0);

                    this.dtMiroFatura.Text = dtGridView1.Rows[numero].Cells[18].Value.ToString();
                    txtNfeMiro.Text = dtGridView1.Rows[numero].Cells[19].Value.ToString();
                    txtVlMiro.Text = dtGridView1.Rows[numero].Cells[30].Value.ToString();
                    this.dtVencimentoMiro.Text = dtGridView1.Rows[numero].Cells[23].Value.ToString();
                    txtPedido.Text = dtGridView1.Rows[numero].Cells[26].Value.ToString();
                    txtCodUnic.Text = dtGridView1.Rows[numero].Cells[25].Value.ToString();
                    txtMiro.Text = dtGridView1.Rows[numero].Cells[29].Value.ToString();
                    txtNf.Text = dtGridView1.Rows[numero].Cells[2].Value.ToString();

                    try
                    {
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT")).Text = this.dtMiroFatura.Text;
                        guiWindow.SendVKey(0);
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR")).Text = txtNfeMiro.Text;
                        guiWindow.SendVKey(0);
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR")).Text = txtVlMiro.Text.Replace(".", ",");
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT")).Text = txtNf.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).Text = txtPedido.Text.Replace(" ", "");
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN")).CaretPosition = 10;
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT")).Text = this.dtVencimentoMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZLSCH")).Text = formPag.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HBKID")).Text = bancoEmpresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HKTID")).Text = bancoEmpresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).Text = txtRefPagmto.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).CaretPosition = 8;
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER")).Text = empresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).Text = txtCodUnic.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE")).Text = "S1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).CaretPosition = 9;
                        guiWindow.SendVKey(0);
                        ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[21]")).Press();
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).Text = txtMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).CaretPosition = 16;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(3);

                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                        GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                        string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));
                        //MessageBox.Show(resultado.Replace(" ", "").Replace("f", "").Replace("o", ""));
                        MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET `miro`='"+ resultado.Replace(" fo","") +"' WHERE nfe='" + txtNfeMiro.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND miro=''", CONEX);
                        DataTable SS = new DataTable();
                        ADAP.Fill(SS);
                        dtGridView1.DataSource = SS;
                        CONEX.Close();
                        numero++;
                    }
                    catch (Exception Err)
                    {
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZFBDT")).Text = this.dtVencimentoMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-ZLSCH")).Text = formPag.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HBKID")).Text = bancoEmpresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-HKTID")).Text = bancoEmpresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).Text = txtRefPagmto.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-KIDNO")).CaretPosition = 8;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER")).Text = empresa.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).Text = txtCodUnic.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE")).Text = "S1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).SetFocus();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/txtINVFO-ZUONR")).CaretPosition = 9;
                        guiWindow.SendVKey(0);
                        ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[21]")).Press();
                        guiWindow.SendVKey(0);
                        ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4")).Select();
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).Text = txtMiro.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).CaretPosition = 16;
                        guiWindow.SendVKey(0);
                        guiWindow.SendVKey(3);

                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();

                        GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");

                        string resultado = statusbar.Text.Substring(13, statusbar.Text.IndexOf('5'));
                        //MessageBox.Show(resultado.Replace(" ", "").Replace("f", "").Replace("o", ""));
                        MySqlConnection CONEX = new MySqlConnection(@"server='"+ txtHost.Text +"';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                        CONEX.Open();
                        MySqlCommand cmd = new MySqlCommand("UPDATE `tb_boleto` SET `miro`='"+ resultado.Replace(" ","").Replace("f","").Replace("o","") +"' WHERE nfe='" + txtNfeMiro.Text + "'", CONEX);
                        cmd.ExecuteNonQuery();
                        MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NOT NULL AND miro=''", CONEX);
                        DataTable SS = new DataTable();
                        ADAP.Fill(SS);
                        dtGridView1.DataSource = SS;
                        CONEX.Close();
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
                MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NULL AND err_col = '1'", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView1.DataSource = SS;
            }
            catch
            {
                MetroMessageBox.Show(this,"Não Existe Itens para Criar Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            try
            {
                MySqlConnection CONEX = new MySqlConnection(@"server='" + txtHost.Text + "';database=sistemagsg_ceal;Uid=energia;Pwd=02984646#Lua;SslMode=none;");
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM tb_boleto WHERE pedido IS NULL AND err_col = '2'", CONEX);
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dtGridView2.DataSource = SS;
            }
            catch
            {
                MetroMessageBox.Show(this,"Não Existe Itens para Criar Pedido!", "Aviso!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
    }
}
