using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;

namespace SistemaGSG
{
    public partial class frmXML : MetroFramework.Forms.MetroForm
    {
        public frmXML()
        {
            InitializeComponent();
        }

        public string vbCr { get; private set; }

        private void frmXML_Load(object sender, EventArgs e)
        {
        }
        
        private void btnImport_Click(object sender, EventArgs e)
        {
            if (MySQL==0)
            {
                int countg = PRODUTOGRID.RowCount;
                int p = 0;
                string Part = "AA";
                try
                {
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
                    //Maximisa Janela
                    guiWindow.Maximize();
                    //Abre Transação
                    Session.SendCommand("/NJ1B1N");
                    //Inicia a Barra de Progresso em 25%
                    LblStatus.Text = "Conexão bem sucedida.......";
                    ProgBar.Maximum = countg;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-NFTYPE")).Text = "I7";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-BUKRS")).Text = "usga";
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-BRANCH")).Text = "0001";
                    ((GuiComboBox)Session.FindById("wnd[0]/usr/cmbJ_1BDYDOC-PARVW")).Key = "LF";
                    ((GuiComboBox)Session.FindById("wnd[0]/usr/cmbJ_1BDYDOC-PARVW")).SetFocus();
                    LblStatus.Text = "Criando Doc. da Nota Fiscal.....";
                    ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-PARID")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-PARID")).CaretPosition = 0;
                    ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(4);
                    ((GuiTextField)Session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).Text = txtCNPJ.Text;
                    ((GuiTextField)Session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).CaretPosition = 14;
                    LblStatus.Text = "Fornecedor Encontrado.....  " + txtEMPRESA.Text + "";
                    ((GuiFrameWindow)Session.FindById("wnd[1]")).SendVKey(0);
                    ((GuiFrameWindow)Session.FindById("wnd[1]")).SendVKey(0);
                    ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                    ((GuiTextField)Session.FindById("wnd[0]/usr/subNF_NUMBER:SAPLJ1BB2:2002/txtJ_1BDYDOC-NFENUM")).Text = txtNFE.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/txtJ_1BDYDOC-SERIES")).Text = txtserie.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-DOCDAT")).Text = dateTimePicker1.Text;
                    ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8")).Select();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_DOCNUM9_DIVIDED-DOCNUM8")).Text = txtNumAleat.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_ACTIVE-CDV")).Text = txtUltDig.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/txtJ_1BDYDOC-AUTHCOD")).Text = txtProt.Text;
                    ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_ACTIVE-CDV")).SetFocus();
                    ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_ACTIVE-CDV")).CaretPosition = 1;
                    ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1")).Select();
                    LblStatus.Text = "Inserindo Itens.....";
                    int Barra = 1;
                    int Enterr = 0;
                    int i = 0;
                    int Position = 9;
                    while (i < countg)
                    {
                        ProgBar.Value = Barra;
                        txtCFOPInic.Text = PRODUTOGRID.Rows[i].Cells["CFOP"].Value.ToString().Trim().Substring(0, 1);
                        txtCFOPUlt.Text = PRODUTOGRID.Rows[i].Cells["CFOP"].Value.ToString().Trim().Substring(1, 3);
                        string CFOP = txtCFOPUlt.Text;
                        txtCountTexto.Text = PRODUTOGRID.Rows[i].Cells["xProd"].Value.ToString().Trim();
                        int ContarTextoProd = txtCountTexto.Text.Length;
                        if (ContarTextoProd > 40)
                        {
                            txtProduto.Text = PRODUTOGRID.Rows[i].Cells["xProd"].Value.ToString().Trim().Substring(0, 40);
                        }
                        else
                        {
                            txtProduto.Text = PRODUTOGRID.Rows[i].Cells["xProd"].Value.ToString().Trim();
                        }
                        decimal Quantidade = 0;
                        Quantidade = Convert.ToDecimal(PRODUTOGRID.Rows[i].Cells["qCom"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                        txtQTD.Text = Quantidade.ToString();
                        decimal ValorUnit = 0;
                        ValorUnit = Convert.ToDecimal(PRODUTOGRID.Rows[i].Cells["vUnCom"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                        textValor1.Text = ValorUnit.ToString();
                        txtUnMedida.Text = PRODUTOGRID.Rows[i].Cells["uCom"].Value.ToString().Trim();
                        if (txtUnMedida.Text == "PEC")
                        {
                            txtUnMedida.Text = "PC";
                        }
                        else
                        {
                            if (txtUnMedida.Text == "qt")
                            {
                                txtUnMedida.Text = "UN";
                            }
                            else
                            {
                                if (txtUnMedida.Text == "UND")
                                {
                                    txtUnMedida.Text = "UN";
                                }
                                else
                                {
                                    txtUnMedida.Text = "UN";
                                }
                            }
                        }
                        txtNCM.Text = PRODUTOGRID.Rows[i].Cells["NCM"].Value.ToString().Trim();
                        if (txtCFOPInic.Text == "5")
                        {
                            string Ini = "1";
                            txtCfop.Text = Ini + CFOP + "/" + Part;
                        }
                        if (txtCFOPInic.Text == "6")
                        {
                            string Ini = "2";
                            txtCfop.Text = Ini + CFOP + "/" + Part;
                        }
                        if (txtCFOPInic.Text == "1")
                        {
                            string Ini = "1";
                            txtCfop.Text = Ini + CFOP + "/" + Part;
                        }
                        if (txtCFOPInic.Text == "2")
                        {
                            string Ini = "2";
                            txtCfop.Text = Ini + CFOP + "/" + Part;
                        }
                        ((GuiCTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-ITMTYP[1," + p + "]")).Text = "1";
                        ((GuiCTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-WERKS[6," + p + "]")).Text = "USGA";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-MAKTX[8," + p + "]")).Text = txtProduto.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATKL[9," + p + "]")).Text = "9010";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-MENGE[10," + p + "]")).Text = txtQTD.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MEINS[11," + p + "]")).Text = txtUnMedida.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-NETPR[19," + p + "]")).Text = textValor1.Text;
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[27," + p + "]")).Text = txtCfop.Text;
                        string IMPOSTOVL = PRODUTOGRID.Rows[i].Cells["IMPOSTO"].Value.ToString().Trim();
                        if (IMPOSTOVL == "0")
                        {
                            ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW1[28," + p + "]")).Text = "IC4";
                        }
                        else
                        {
                            ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW1[28," + p + "]")).Text = "IC0";
                        }
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW2[29," + p + "]")).Text = "IP1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW4[31," + p + "]")).Text = "C70";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW5[32," + p + "]")).Text = "P70";
                        ((GuiComboBox)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/cmbJ_1BDYLIN-MATORG[33," + p + "]")).Key = "0";
                        ((GuiComboBox)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/cmbJ_1BDYLIN-MATUSE[34," + p + "]")).Key = "1";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-NBM[35," + p + "]")).Text = txtNCM.Text.Replace(',', '.');
                        p++;
                        i++;
                        LblStatus.Text = "Inserindo Itens.....";
                        Barra++;
                        
                        if (p == 10)
                        {
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            ((GuiTableControl)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL")).VerticalScrollbar.Position = Position;
                            p = 1;
                            Position += 9;
                        }
                    }
                    
                    LblStatus.Text = "Itens Inseridos com Sucesso.....";
                    if (i == countg)
                    {
                        int Enter = 0;
                        int CountgEnt = countg;
                        CountgEnt++;
                        int ContItens = CountgEnt;
                        while (Enter < ContItens)
                        {
                            ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(0);
                            Enter++;
                        }
                        ((GuiButton)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/btn%#AUTOTEXT004")).Press();
                        ((GuiButton)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/btn%#AUTOTEXT002")).Press();
                        int AddImpost = 0;
                        int ValTotItem = 0;

                        LblStatus.Text = "Inserindo códigos de Imposto.....";
                        int CountImposto = 0;

                        while (AddImpost < countg)
                        {
                            decimal ValorTotal = 0;
                            ValorTotal = Convert.ToDecimal(PRODUTOGRID.Rows[ValTotItem].Cells["vProd"].Value.ToString().Replace('.', ',').Trim().Substring(0, 4));
                            textValor1.Text = ValorTotal.ToString();

                            int ContGridImposto = ICMS00GRID.RowCount;
                            decimal ImpostoValor = 0;
                            decimal ImpostoBaseValor = 0;
                            decimal ImpostoValorICMS = 0;

                            ImpostoValor = Convert.ToDecimal(ICMS00GRID.Rows[CountImposto].Cells["pICMS"].Value.ToString().Replace('.', ',').Trim().Substring(0, 4));
                            txtImposto.Text = ImpostoValor.ToString();
                            ImpostoBaseValor = Convert.ToDecimal(ICMS00GRID.Rows[CountImposto].Cells["vBC"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                            txtBaseCalc.Text = ImpostoBaseValor.ToString();
                            ImpostoValorICMS = Convert.ToDecimal(ICMS00GRID.Rows[CountImposto].Cells["vICMS"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                            txtValorICMS.Text = ImpostoValorICMS.ToString();

                            if (txtImposto.Text == "0")
                            {
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,0]")).Text = "ICM0";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,1]")).Text = "IPI0";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-EXCBAS[6,0]")).Text = textValor1.Text;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-EXCBAS[6,1]")).Text = textValor1.Text;
                            }
                            else
                            {
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,0]")).Text = "ICM2";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-RATE[4,0]")).Text = txtImposto.Text;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-TAXVAL[5,0]")).Text = txtValorICMS.Text;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-OTHBAS[7,0]")).Text = txtBaseCalc.Text;
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,1]")).Text = "IPI0";
                                ((GuiTextField)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-EXCBAS[6,1]")).Text = textValor1.Text;
                            }
                            CountImposto++;
                            AddImpost++;
                            ValTotItem++;
                            ((GuiButton)Session.FindById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/btnPB_NEXT")).Press();
                        }
                        if (countg == 1)
                        {
                            Barra = 60;
                        }
                        else
                        {
                            Barra++;
                        }
                        ProgBar.Value = Barra;
                        ((GuiMenu)Session.FindById("wnd[0]/mbar/menu[2]/menu[3]/menu[0]")).Select();
                        ((GuiTableControl)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL")).Columns.ElementAt(0);
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).Text = "ENTRADA PARA REGULARIZAÇÃO.";
                        ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]")).CaretPosition = 29;
                        ((GuiFrameWindow)Session.FindById("wnd[0]")).SendVKey(2);
                        LblStatus.Text = "Adicionando informações adicionais no campo de mensagem.....";

                        ((GuiShell)Session.FindById("wnd[0]/usr/cntlEDITOR/shellcont/shell")).Text = "ENTRADA PARA REGULARIZAÇÃO." + vbCr + "  " + txtDados.Text + vbCr + "";
                        ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                        ((GuiButton)Session.FindById("wnd[1]/usr/btnBUTTON_1")).Press();
                        LblStatus.Text = "Verifique se os Valores estão corretos.....";

                        if (MessageBox.Show("Valores Corretos?!", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            //Pressiona o Botão Gravar
                            //((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[11]")).Press();
                            //Pega a Barra de Status do SAP
                            GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                            //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                            string resultado = statusbar.Text.Substring(12, 10);
                            try
                            {

                                MySqlCommand cmd = new MySqlCommand("INSERT INTO `tb_result` (`col_notafiscal`, `col_DocNumSAP`) VALUES ('" + txtNFE.Text.Trim() + "', '" + resultado.Trim() + "')", ConexaoDados.GetConnectionXML());
                                cmd.ExecuteNonQuery();
                                ConexaoDados.GetConnectionXML().Close();
                            }
                            catch (MySqlException ErroMysql)
                            {
                                MessageBox.Show(ErroMysql.Message);
                            }
                            if (countg == 1)
                            {
                                Barra = 100;
                            }
                            else
                            {
                                Barra = 100;
                            }
                            ProgBar.Value = Barra;
                            LblStatus.Text = "Comcuído com Sucesso..... " + resultado.Trim() + " ";
                            Clear();
                        }
                        else
                        {
                            if (MessageBox.Show("Nota Fiscal Concluída?!", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                //Pega a Barra de Status do SAP
                                GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
                                //Me retorna apenas o número do pedido no tratamento da importação no Banco de Dados ele retira o º e os espaços.
                                string resultado = statusbar.Text.Substring(12, 10);
                                try
                                {

                                    MySqlCommand cmd = new MySqlCommand("INSERT INTO `tb_result` (`col_notafiscal`, `col_DocNumSAP`) VALUES ('" + txtNFE.Text.Trim() + "', '" + resultado.Trim() + "')", ConexaoDados.GetConnectionXML());
                                    cmd.ExecuteNonQuery();
                                    ConexaoDados.GetConnectionXML().Close();
                                }
                                catch (MySqlException ErroMysql)
                                {
                                    MessageBox.Show(ErroMysql.Message);
                                }
                                if (countg == 1)
                                {
                                    Barra = 100;
                                }
                                else
                                {
                                    Barra = 100;
                                }
                                ProgBar.Value = Barra;
                                LblStatus.Text = "Comcuído com Sucesso..... " + resultado.Trim() + " ";
                                Clear();
                            }
                        }
                    }
                }
                catch (Exception Err)
                {
                    MessageBox.Show(Err.Message);
                }
            }
            else
            {
                try
                {
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = "yyyy-MM-dd";
                    MySqlCommand prompt_cmd = new MySqlCommand("UPDATE `tb_chave` SET empresa='" + txtEMPRESA.Text.Trim() + "' ,n_nfe='" + txtNFE.Text.Trim() + "',emisao='" + this.dateTimePicker1.Text + "' WHERE col_chave='" + txtCHAVE.Text + "'", ConexaoDados.GetConnectionXML());
                    prompt_cmd.ExecuteNonQuery();
                    ConexaoDados.GetConnectionXML().Close();
                    dateTimePicker1.Format = DateTimePickerFormat.Custom;
                    dateTimePicker1.CustomFormat = "dd.MM.yyyy";
                    MessageBox.Show("Sucesso!");
                }
                catch(MySqlException ErrMy)
                {
                    MessageBox.Show("Erro no Banco de Dados! - /n"+ErrMy.Message);
                } 
                catch (Exception Err)
                {
                    MessageBox.Show(Err.Message);
                }
            }
        }
        private void Clear()
        {
            txtProduto.Text = "";
            txtQTD.Text = "";
            textValor1.Text = "";
            txtUnMedida.Text = "";
            txtCfop.Text = "";
            txtNCM.Text = "";
            txtValorICMS.Text = "";
            txtImposto.Text = "";
            txtBaseCalc.Text = "";
        }
        private void btnAbrirXML_Click_1(object sender, EventArgs e)
        {
            MySQL = 0;
            string VALORIMPOSTO = "12";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Arquivos XML(*.xml)|*.xml";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string xmlFilePath = openFileDialog.FileName;
                using (DataSet ds = new DataSet())
                {
                    try
                    {
                        ds.ReadXml(xmlFilePath);
                        CNPJGRID.DataSource = ds.Tables["emit"];
                        txtCNPJ.Text = CNPJGRID.Rows[0].Cells["CNPJ"].Value.ToString().Trim();
                        txtEMPRESA.Text = CNPJGRID.Rows[0].Cells["xNome"].Value.ToString().Trim();
                        //----------------
                        NFEGRID.DataSource = ds.Tables["ide"];
                        txtNFE.Text = NFEGRID.Rows[0].Cells["nNF"].Value.ToString().Trim();
                        txtserie.Text = NFEGRID.Rows[0].Cells["serie"].Value.ToString().Trim();
                        txtDate.Text = NFEGRID.Rows[0].Cells["dhEmi"].Value.ToString().Trim().Substring(0, 10);
                        dateTimePicker1.Text = txtDate.Text;
                        //----------------
                        DADOSGRID.DataSource = ds.Tables["infAdic"];
                        txtDados.Text = DADOSGRID.Rows[0].Cells["infCpl"].Value.ToString().Trim();
                        //----------------
                        CHAVEGRID.DataSource = ds.Tables["infProt"];
                        txtCHAVE.Text = CHAVEGRID.Rows[0].Cells["chNFe"].Value.ToString().Trim();
                        txtProt.Text = CHAVEGRID.Rows[0].Cells["nProt"].Value.ToString().Trim();
                        txtNumAleat.Text = txtCHAVE.Text.Substring(35, 8);
                        txtUltDig.Text = txtCHAVE.Text.Substring(43, 1);
                        //----------------
                        PRODUTOGRID.DataSource = ds.Tables["prod"];
                        foreach(DataGridViewRow MyDgr in PRODUTOGRID.Rows)
                        {
                            MyDgr.Cells["IMPOSTO"].Value = VALORIMPOSTO;
                        }
                        //----------------
                        ICMS00GRID.DataSource = ds.Tables["ICMS00"];
                        int ContGridImposto = ICMS00GRID.RowCount;
                        if (ContGridImposto == 0)
                        {

                        }
                        else
                        {
                            decimal ImpostoValor = 0;
                            ImpostoValor = Convert.ToDecimal(ICMS00GRID.Rows[0].Cells["pICMS"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                            txtImposto.Text = ImpostoValor.ToString();
                            decimal ImpostoBaseValor = 0;
                            ImpostoBaseValor = Convert.ToDecimal(ICMS00GRID.Rows[0].Cells["vBC"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                            txtBaseCalc.Text = ImpostoBaseValor.ToString();
                            decimal ImpostoValorICMS = 0;
                            ImpostoValorICMS = Convert.ToDecimal(ICMS00GRID.Rows[0].Cells["vICMS"].Value.ToString().Replace('.', ',').TrimEnd('0'));
                            txtValorICMS.Text = ImpostoValorICMS.ToString();
                        }
                        ICMS30GRID.DataSource = ds.Tables["ICMS30"];
                        IMPOSTOGRID.DataSource = ds.Tables["imposto"];
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.Message);
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main back = new frm_Main();
                back.Show();
                Close();
            }
        }
        int MySQL = 0;
        private void button2_Click(object sender, EventArgs e)
        {
            

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (MySQL == 0)
            {
                MySQL++;
            }
            else
            {

            }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string xmlFilePath = openFileDialog.FileName;
                using (DataSet ds = new DataSet())
                {
                    try
                    {
                        ds.ReadXml(xmlFilePath);
                        CNPJGRID.DataSource = ds.Tables["emit"];
                        txtCNPJ.Text = CNPJGRID.Rows[0].Cells["CNPJ"].Value.ToString().Trim();
                        txtEMPRESA.Text = CNPJGRID.Rows[0].Cells["xNome"].Value.ToString().Trim();
                        //----------------
                        NFEGRID.DataSource = ds.Tables["ide"];
                        txtNFE.Text = NFEGRID.Rows[0].Cells["nNF"].Value.ToString().Trim();
                        txtserie.Text = NFEGRID.Rows[0].Cells["serie"].Value.ToString().Trim();
                        txtDate.Text = NFEGRID.Rows[0].Cells["dhEmi"].Value.ToString().Trim().Substring(0, 10);
                        dateTimePicker1.Text = txtDate.Text;
                        //----------------
                        DADOSGRID.DataSource = ds.Tables["infAdic"];
                        txtDados.Text = DADOSGRID.Rows[0].Cells["infCpl"].Value.ToString().Trim();
                        //----------------
                        CHAVEGRID.DataSource = ds.Tables["infProt"];
                        txtCHAVE.Text = CHAVEGRID.Rows[0].Cells["chNFe"].Value.ToString().Trim();
                        txtProt.Text = CHAVEGRID.Rows[0].Cells["nProt"].Value.ToString().Trim();
                        txtNumAleat.Text = txtCHAVE.Text.Substring(35, 8);
                        txtUltDig.Text = txtCHAVE.Text.Substring(43, 1);
                        //----------------
                        PRODUTOGRID.DataSource = ds.Tables["prod"];
                        //----------------
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.Message);
                    }
                }
            }
        }
    }
}
