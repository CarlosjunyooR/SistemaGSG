using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SAPFEWSELib;
using SapROTWr;
using MySql.Data.MySqlClient;

namespace SistemaGSG
{
    public partial class FormAcesso : MetroFramework.Forms.MetroForm
    {
        public FormAcesso()
        {
            InitializeComponent();
            lblUnd.Text = "SACOS";
            UltimNumSacaria();
        }
        String URL1 = "http://localhost/sistemagsgv2.0/template/dashboard/pages/relatorios/faturamento/acesso/RelatorioAcesso.php";
        String URL2 = "http://localhost/sistemagsgv2.0/template/dashboard/pages/relatorios/faturamento/acesso/RelatorioCheckList.php";
        string usuarioLogado = dados.usuario;
        private void UltimNumSacaria()
        {

            MySqlCommand MyCommand = new MySqlCommand();
            MyCommand.Connection = ConexaoDados.GetConnectionFaturameto();
            MyCommand.CommandText = "SELECT * FROM tb_acesso ORDER BY id DESC";
            MySqlDataReader dreader = MyCommand.ExecuteReader();
            while (dreader.Read())
            {
                txtSacariaInic.Text = dreader["col_sacaria_fim"].ToString();
                break;
            }
            ConexaoDados.GetConnectionFaturameto().Close();
            double SomaUm = Convert.ToDouble(txtSacariaInic.Text.Trim()) + 1;
            txtSacariaInic.Text = SomaUm.ToString();
        }
        private void AbrirPDFs()
        {
            var Urk = URL1;
            var Urb = URL2;

            var AbrirNavegador1 = new Navegador(Urk.ToString());
            AbrirNavegador1.Show();

            var AbrirNavegador2 = new Navegador(Urb.ToString());
            AbrirNavegador2.Show();
        }
        private void Test()
        {
            // The input string.
            string s = "I have a cat";

            // Loop through all instances of the letter a.
            int i = 0;
            while ((i = s.IndexOf('a', i)) != -1)
            {
                // Print out the substring.
                //Console.WriteLine(s.Substring(i));
                MessageBox.Show(s.Substring(i));

                // Increment the index.
                i++;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Test();
        }
        private void SAPAcesso()
        {
            ProgressBar.Value = 0;

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
            UserTXT.Text = Session.Info.User;
            //Pega a referência ao "FRAME" principal para enviar comandos de chaves virtuais o SAP
            GuiFrameWindow guiWindow = Session.ActiveWindow;
            //Abre Transação
            Session.SendCommand("/nxk03");
            guiWindow.SendVKey(4);
            ProgressBar.Value = 10;
            //Pega a Barra de Status do SAP
            ((GuiTextField)Session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]")).Text = txtCPF.Text;
            ((GuiFrameWindow)Session.FindById("wnd[1]")).SendVKey(0);

            //try
            //{
            //    GuiStatusbar statusbar = (GuiStatusbar)Session.FindById("wnd[0]/sbar");
            //    string resultado = statusbar.Text.Substring(0, 30);
            //    if(resultado == "Nenhum valor para esta seleção")
            //    {
            //        MessageBox.Show("Motorista não cadastrado!");
            //    }
            //}
            //catch(Exception Err)
            //{
            //    MessageBox.Show(Err.Message);
            //}Quantidade informada excede capacidade de carregamento.
            //System.Diagnostics.Debugger.Break();

            String NMotor = ((GuiLabel)Session.FindById("wnd[1]/usr/lbl[93,3]")).Text;
            guiWindow.SendVKey(0);
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS")).Text = "";
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0110")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0120")).Selected = true;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0130")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkWRF02K-D0380")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0210")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0215")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0220")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0610")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0310")).Selected = false;
            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkWRF02K-D0320")).Selected = false;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-EKORG")).Text = "1000";
            guiWindow.SendVKey(0);
            String CMotor = ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR")).Text;
            //txtCPFCompleto.Text = CMotor;
            txtCPF.Visible = false;
            txtCPFCompleto.Visible = true;
            txtCPFCompleto.Text = NMotor;
            try
            {
                //Abrir Conexão.
                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_motorista WHERE col_cod_id ='" + CMotor + "' ", ConexaoDados.GetConnectionFaturameto());//Seleção da tabela no Banco de Dados.
                //Executa o comando.
                prompt.ExecuteNonQuery();
                //Converte o resultado para números inteiros.
                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                //Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                if (consultDB > 0)
                {

                }
                else
                {
                    try
                    {
                        MySqlCommand command = new MySqlCommand("INSERT INTO `tb_motorista` (`col_cod_id`, `col_nome_mot`) VALUES ('" + CMotor + "', '" + NMotor + "')", ConexaoDados.GetConnectionFaturameto());
                        command.ExecuteNonQuery();
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.Message);
                        ProgressBar.Value = 0;
                    }
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Olá Srº(a), " + usuarioLogado + " contate o administrador, houve algum erro na\naplicação!.");
            }
            finally
            {

            }
            ProgressBar.Value = 20;

            String CTransportador = ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtLFA1-FISKU")).Text;
            if (String.IsNullOrEmpty(CTransportador))
            {

            }
            else
            {
                Session.SendCommand("/nxk03");
                ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS")).Text = " ";
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0110")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0120")).Selected = true;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0130")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkWRF02K-D0380")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0210")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0215")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0220")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0610")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkRF02K-D0310")).Selected = false;
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkWRF02K-D0320")).Selected = false;
                ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR")).Text = CTransportador.Trim();
                ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtRF02K-EKORG")).Text = "1000";
                guiWindow.SendVKey(0);
                String NTransportador = ((GuiTextField)Session.FindById("wnd[0]/usr/txtLFA1_INT-NAME1")).Text;
                String CNPJTransportador = ((GuiTextField)Session.FindById("wnd[0]/usr/txtLFA1-STCD1")).Text;
                String IEstTransportador = ((GuiTextField)Session.FindById("wnd[0]/usr/txtLFA1-STCD3")).Text;
                maskTransp.Text = CNPJTransportador;
                maskCliente.Text = txtCliente.Text;
                txtTransportadora.Text = NTransportador;
                txtIEST.Text = IEstTransportador;
            }
            ProgressBar.Value = 30;

            Session.SendCommand("/nxd03");
            guiWindow.SendVKey(4);
            ((GuiTextField)Session.FindById("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]")).Text = txtCliente.Text;
            ((GuiFrameWindow)Session.FindById("wnd[2]")).SendVKey(0);
            ((GuiFrameWindow)Session.FindById("wnd[2]")).SendVKey(0);
            ((GuiFrameWindow)Session.FindById("wnd[1]")).SendVKey(0);
            String NCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1")).Text;
            String RuaCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET")).Text;
            String CidCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1")).Text;
            String EstCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION")).Text;
            ((GuiTab)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02")).Select();
            String IesCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/txtKNA1-STCD3")).Text;
            String CCliente = ((GuiTextField)Session.FindById("wnd[0]/usr/subSUBKOPF:SAPMF02D:7001/ctxtRF02D-KUNNR")).Text;
            try
            {

                //Seleção da tabela no Banco de Dados.
                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_cliente WHERE col_cod_id ='" + CCliente + "' ", ConexaoDados.GetConnectionFaturameto());
                //Executa o comando.
                prompt.ExecuteNonQuery();
                //Converte o resultado para números inteiros.
                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                //Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                if (consultDB > 0)
                {

                }
                else
                {
                    try
                    {
                        MySqlCommand command = new MySqlCommand("INSERT INTO `tb_cliente` (`col_cod_id`, `col_nome`, `col_cnpj`, `col_rua`, `col_cidade`, `col_iestadual`, `col_estado`) VALUES ('" + CCliente + "', '" + NCliente + "', '" + maskCliente.Text + "', '" + RuaCliente + "', '" + CidCliente + "', '" + IesCliente + "', '" + EstCliente + "')", ConexaoDados.GetConnectionFaturameto());
                        command.ExecuteNonQuery();
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.Message);
                    }
                }

            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Olá Srº(a), " + usuarioLogado + " contate o administrador, houve algum erro na\naplicação!.");
                ProgressBar.Value = 0;
            }
            finally
            {

            }
            Session.SendCommand("/NZBL014");
            ProgressBar.Value = 50;

            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_1000-LIFNR")).Text = CMotor;
            if (String.IsNullOrEmpty(CTransportador))
            {
                ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkT_1000-AUTO")).Selected = true;
                guiWindow.SendVKey(0);
            }
            else
            {
                ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_1000-LIFNR2")).Text = CTransportador;
                guiWindow.SendVKey(0);
                try
                {
                    //Seleção da tabela no Banco de Dados.
                    MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_transport WHERE col_cod_id ='" + CTransportador + "' ", ConexaoDados.GetConnectionFaturameto());
                    //Executa o comando.
                    prompt.ExecuteNonQuery();
                    //Converte o resultado para números inteiros.
                    int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                    //Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                    if (consultDB > 0)
                    {

                    }
                    else
                    {
                        try
                        {
                            MySqlCommand command = new MySqlCommand("INSERT INTO `tb_transport` (`col_cod_id`, `col_nome_transp`, `col_cnpj_transp`, `col_iestadual_transp`) VALUES ('" + CTransportador + "', '" + txtTransportadora.Text + "', '" + maskTransp.Text + "', '" + txtIEST.Text + "')", ConexaoDados.GetConnectionFaturameto());
                            command.ExecuteNonQuery();
                            ProgressBar.Value = 65;
                        }
                        catch (Exception Err)
                        {
                            MessageBox.Show(Err.Message);
                        }
                    }
                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Olá Srº(a), " + usuarioLogado + " contate o administrador, houve algum erro na\naplicação!.");
                    ProgressBar.Value = 0;
                }
                finally
                {

                }
            }
            ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[18]")).Press();
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_2000-PLACA1")).Text = txtPlaca1.Text;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_2000-PLACA2")).Text = txtPlaca2.Text;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_2000-PLACA3")).Text = txtPlaca3.Text;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT_2000-PLACA4")).Text = txtPlaca4.Text;
            guiWindow.SendVKey(0);
            ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[18]")).Press();
            guiWindow.SendVKey(0);
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtT001W-BWKEY")).Text = "USGA";
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtVBAK-KUNNR")).Text = CCliente;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtMARA-MATNR")).Text = "100000";
            ((GuiTextField)Session.FindById("wnd[0]/usr/txtZMM_STAT_VEHI-QUANTID")).Text = txtQuantidade.Text;
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtVBAK-KUNNR")).SetFocus();
            ((GuiCTextField)Session.FindById("wnd[0]/usr/ctxtVBAK-KUNNR")).CaretPosition = 10;
            guiWindow.SendVKey(0);
            guiWindow.SendVKey(0);

            ((GuiCheckBox)Session.FindById("wnd[0]/usr/chkCK_FAT_ACESSO")).Selected = true;
            ((GuiTextField)Session.FindById("wnd[0]/usr/tblSAPMZMM099TC_DISP/txtTVBAP-CODIF[4,0]")).Text = ""+ txtSacariaInic.Text +" a "+ txtSacariaFim.Text +"";
            try
            {
                if (String.IsNullOrEmpty(CTransportador))
                {
                    MySqlCommand command = new MySqlCommand("INSERT INTO `tb_acesso` (`col_acesso`, `col_data_acesso`, `col_hora_acesso`, `col_cliente`, `col_tipo_sac`, `col_quant`, `col_ov`, `col_sacaria_inicio`, `col_sacaria_fim`, `col_safra`, `col_transport`, `col_motorista`, `col_placa_1`, `col_placa_2`, `col_placa_3`, `col_placa_4`, `Obs`) VALUES ('0000125371', CURDATE(), NOW(), '" + CCliente + "', '" + lblUnd.Text + "', '" + txtQuantidade.Text + "', '" + txtOv.Text + "', '" + txtSacariaInic.Text + "', '" + txtSacariaFim.Text + "', '" + maskSafra.Text + "', NULL, '" + CMotor + "', '" + txtPlaca1.Text + "', '" + txtPlaca2.Text + "', '" + txtPlaca3.Text + "', '" + txtPlaca4.Text + "', '" + richObs.Text + "')", ConexaoDados.GetConnectionFaturameto());
                    command.ExecuteNonQuery();
                    ProgressBar.Value = 70;
                }
                else
                {
                    MySqlCommand command = new MySqlCommand("INSERT INTO `tb_acesso` (`col_acesso`, `col_data_acesso`, `col_hora_acesso`, `col_cliente`, `col_tipo_sac`, `col_quant`, `col_ov`, `col_sacaria_inicio`, `col_sacaria_fim`, `col_safra`, `col_transport`, `col_motorista`, `col_placa_1`, `col_placa_2`, `col_placa_3`, `col_placa_4`, `Obs`) VALUES ('0000125371', CURDATE(), NOW(), '" + CCliente + "', '" + lblUnd.Text + "', '" + txtQuantidade.Text + "', '" + txtOv.Text + "', '" + txtSacariaInic.Text + "', '" + txtSacariaFim.Text + "', '" + maskSafra.Text + "', '" + CTransportador + "', '" + CMotor + "', '" + txtPlaca1.Text + "', '" + txtPlaca2.Text + "', '" + txtPlaca3.Text + "', '" + txtPlaca4.Text + "', '" + richObs.Text + "')", ConexaoDados.GetConnectionFaturameto());
                    command.ExecuteNonQuery();
                    ProgressBar.Value = 70;
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Olá Srº(a), " + usuarioLogado + " contate o administrador, houve algum erro na\naplicação!.");
                ProgressBar.Value = 0;
            }
            catch (MySqlException eer)
            {
                MessageBox.Show(eer.Message);
            }
            finally
            {

            }
            ProgressBar.Value = 100;

            //Fecha Conexão
            ConexaoDados.GetConnectionFaturameto().Close();
            Session.SendCommand("/N");
            CTransportador = "";
            AbrirPDFs();
            UltimNumSacaria();
            txtSacariaFim.Text = "";
        }
        private void LimparCampos()
        {
            txtCPF.Visible = true;
            txtCPFCompleto.Visible = false;
            txtCPFCompleto.Text = "";
            txtTransportadora.Text = "";
            txtPlaca1.Text = "";
            txtPlaca2.Text = "";
            txtPlaca3.Text = "";
            txtPlaca4.Text = "";
            txtQuantidade.Text = "";
            txtCliente.Text = "";
            richObs.Text = "";
            txtOv.Text = "";
            txtCPF.Text = "";
        }
        private void txtCPF_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SAPAcesso();
        }
        private void SomaSacaria()
        {
            double DimSacariaFim = Convert.ToDouble(txtQuantidade.Text.Trim()) - 1;
            double ValorInicio = Convert.ToDouble(txtSacariaInic.Text.Trim());
            double SomaSacariaFim = Convert.ToDouble(DimSacariaFim) + ValorInicio;
            txtSacariaFim.Text = SomaSacariaFim.ToString();
        }
        private void textBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (ckbag.Checked)
            {

            }
            else
            {
                if (e.KeyCode == Keys.Enter)
                {
                    SomaSacaria();
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbag.Checked)
            {
                lblSacaria.Visible = false;
                txtSacariaInic.Visible = false;
                txtSacariaFim.Visible = false;
                txtOv.Location = new Point(54, 94);
                lblOv.Location = new Point(9, 97);
                lblUnd.Text = "BAG";
                txtQuantidade.Text = "22";
                lblBagQuant.Text = "27000";
            }
            else
            {
                lblSacaria.Visible = true;
                txtSacariaInic.Visible = true;
                txtSacariaFim.Visible = true;
                txtOv.Location = new Point(54, 120);
                lblOv.Location = new Point(9, 123);
                lblUnd.Text = "SACOS";
                txtQuantidade.Text = "640";
                lblBagQuant.Text = "";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LimparCampos();
        }

        private void button4_Click(object sender, EventArgs e)
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
