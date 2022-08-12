using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class FormDownloadXML : MetroFramework.Forms.MetroForm
    {
        public FormDownloadXML()
        {
            InitializeComponent();
            VerifyVersion(webBrowser1);
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.Navigate("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=d09fwabTnLk=");
            LoadDataGrid();
            lblPasta.Text = @"C:\ArquivosSAP\XML\";
        }
        private void LoadDataGrid()
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_chave` WHERE col_Downl='1' ORDER BY `tb_chave`.`emisao` DESC", ConexaoDados.GetConnectionXML());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridView1.DataSource = SS;
            ConexaoDados.GetConnectionXML().Close();
            int countg = dataGridView1.RowCount;
            lblResultQuant.Text = countg.ToString();
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

        private void button1_Click(object sender, EventArgs e)
        {
            //Inserir valor no Input Text da Página
            webBrowser1.Document.GetElementById("ctl00_ContentPlaceHolder1_txtChaveAcessoResumo").InnerText = dataGridView1.Rows[0].Cells["col_chave"].Value.ToString();
            try
            {
                MySqlCommand prompt_cmd = new MySqlCommand("UPDATE `tb_chave` SET col_Downl='2' WHERE col_chave='" + dataGridView1.Rows[0].Cells["col_chave"].Value.ToString() + "'", ConexaoDados.GetConnectionXML());
                prompt_cmd.ExecuteNonQuery();
                ConexaoDados.GetConnectionXML().Close();
            }
            catch (MySqlException ErrMy)
            {
                MessageBox.Show("Erro no Banco de Dados! - /n" + ErrMy.Message);
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            LoadDataGrid();
        }

        private void FormDownloadXML_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frmProtocolo frm_Main = new frmProtocolo();
                frm_Main.Show();
                Close();
            }
        }
    }
}
