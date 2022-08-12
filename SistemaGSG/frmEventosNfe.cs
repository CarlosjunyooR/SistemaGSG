using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class frmEventosNfe : MetroFramework.Forms.MetroForm
    {
        public frmEventosNfe()
        {
            InitializeComponent();
            VerifyVersion(WebBrowserSefaz);
            WebBrowserSefaz.ScriptErrorsSuppressed = true;
            WebBrowserSefaz.Navigate("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ%20gAVw2g=");
            //WebBrowserSefaz.Navigate("http://localhost/nota/default2.html");
            LoadDataGrid();
        }
        private void LoadDataGrid()
        {
            if (radioButton3.Checked)
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_nota_consultar` WHERE col_status='1' ORDER BY `tb_nota_consultar`.`col_id` ASC", ConexaoDados.GetConnectionFaturameto());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
                ConexaoDados.GetConnectionFaturameto().Close();
                int countg = dataGridView1.RowCount;
                lblCount.Text = countg.ToString();
            }
            if (radioButton1.Checked)
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_nota_consultar` WHERE col_status='143' ORDER BY `tb_nota_consultar`.`col_id` ASC", ConexaoDados.GetConnectionFaturameto());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
                ConexaoDados.GetConnectionFaturameto().Close();
                int countg = dataGridView1.RowCount;
                lblCount.Text = countg.ToString();
            }
            if (radioButton2.Checked)
            {
                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_nota_consultar` WHERE col_status='142' ORDER BY `tb_nota_consultar`.`col_id` ASC", ConexaoDados.GetConnectionFaturameto());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridView1.DataSource = SS;
                ConexaoDados.GetConnectionFaturameto().Close();
                int countg = dataGridView1.RowCount;
                lblCount.Text = countg.ToString();
            }
        }
        public void VerifyVersion(WebBrowser webbrowser)
        {
            string appName = "";
            try
            {
                appName = System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe";

                RegistryKey fbeKey = null;
                ////For 64 bit Machine 
                fbeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                if (fbeKey == null)
                    fbeKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION");
                using (fbeKey)
                {
                    fbeKey.SetValue(appName, 11000, RegistryValueKind.DWord);
                }


                //For 32 bit Machine 
                fbeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                if (fbeKey == null)
                    fbeKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION");
                using (fbeKey)
                {
                    fbeKey.SetValue(appName, 11000, RegistryValueKind.DWord);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(appName + "\n" + ex.ToString(), "Unexpected error setting browser mode!");
            }
        }

        private void ColetarInfo()
        {
            HtmlElementCollection Pesquisa = this.WebBrowserSefaz.Document.GetElementsByTagName("a");
            try
            {
                int LinhaDG = 0;
                foreach (HtmlElement Funcao in Pesquisa)
                {
                    if (Funcao.GetAttribute("id").Equals("lnkCce"))
                    {
                        string Protocolo;
                        Protocolo = Funcao.GetAttribute("innerText");

                        HtmlElementCollection PesquisaDesc = this.WebBrowserSefaz.Document.GetElementsByTagName("div");
                        foreach (HtmlElement htmlElement in PesquisaDesc)
                        {
                            if (htmlElement.GetAttribute("id").Equals("CienciaOperacao" + Protocolo.Trim() + ""))
                            {
                                
                                string CienciaOperacao;

                                txtNfe.Text = dataGridView1.Rows[LinhaDG].Cells["col_Nfe"].Value.ToString();

                                CienciaOperacao = "Ciência da Operação pelo Destinatário (Órgão Autor: AN)";

                                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(col_Nfe) FROM tb_evento WHERE col_protocolo='" + Protocolo + "' ", ConexaoDados.GetConnectionFaturameto());
                                prompt.ExecuteNonQuery();
                                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                                if (consultDB > 0)
                                {

                                }
                                else
                                {
                                    MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO `tb_evento` (`col_Nfe`, `col_eventoNfe`, `col_protocolo`, `col_dataEvento`) VALUES ('" + txtNfe.Text.Trim() + "', '" + CienciaOperacao + "', '" + Protocolo + "', '2021-10-19')", ConexaoDados.GetConnectionFaturameto());
                                    prompt_cmd.ExecuteNonQuery();
                                    ConexaoDados.GetConnectionFaturameto().Close();
                                }
                            }
                            if (htmlElement.GetAttribute("id").Equals("ConfirmacaoOperacao" + Protocolo.Trim() + ""))
                            {
                                string ConfirmacaoOperacao;
                                ConfirmacaoOperacao = "Confirmação da Operação pelo Destinatário (Órgão Autor: AN)";
                                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(col_Nfe) FROM tb_evento WHERE col_protocolo='" + Protocolo + "' ", ConexaoDados.GetConnectionFaturameto());
                                prompt.ExecuteNonQuery();
                                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                                if (consultDB > 0)
                                {

                                }
                                else
                                {
                                    MySqlCommand prompt_cmd_novo = new MySqlCommand("INSERT INTO `tb_evento` (`col_Nfe`, `col_eventoNfe`, `col_protocolo`, `col_dataEvento`) VALUES ('" + txtNfe.Text + "', '" + ConfirmacaoOperacao + "', '" + Protocolo.Trim() + "', '2021-10-19')", ConexaoDados.GetConnectionFaturameto());
                                    prompt_cmd_novo.ExecuteNonQuery();
                                    ConexaoDados.GetConnectionFaturameto().Close();
                                    try
                                    {
                                        MySqlCommand UpdateCMD = new MySqlCommand("UPDATE `tb_nota_consultar` SET col_status='2' WHERE col_chaveAcesso='" + dataGridView1.Rows[LinhaDG].Cells["col_chaveAcesso"].Value.ToString() + "'", ConexaoDados.GetConnectionFaturameto());
                                        UpdateCMD.ExecuteNonQuery();
                                        ConexaoDados.GetConnectionFaturameto().Close();
                                        LoadDataGrid();
                                    }
                                    catch (Exception Er)
                                    {
                                        MessageBox.Show(Er.Message);
                                    }
                                }
                            }
                        }
                        MessageBox.Show("Finalizado!");
                    }
                }
            }
            catch (Exception Error)
            {
                MessageBox.Show(Error.Message);
            }
        }

        private void btnTeste_Click(object sender, EventArgs e)
        {
            HtmlElementCollection Pesquisa = this.WebBrowserSefaz.Document.GetElementsByTagName("input");
            try
            {
                foreach (HtmlElement Funcao in Pesquisa)
                {
                    if (Funcao.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_btnConsultaCompleta"))
                    {
                        Funcao.InvokeMember("Click");
                    }
                }
                ColetarInfo();
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WebBrowserSefaz.Document.GetElementById("ctl00_ContentPlaceHolder1_txtChaveAcessoResumo").InnerText = dataGridView1.Rows[0].Cells["col_chaveAcesso"].Value.ToString();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main back = new frm_Main();
                back.Show();
                Close();
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            LoadDataGrid();
        }
    }
}
