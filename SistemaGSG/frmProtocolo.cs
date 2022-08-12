using Microsoft.Win32;
using MySql.Data.MySqlClient;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;


namespace SistemaGSG
{
    public partial class frmProtocolo : MetroFramework.Forms.MetroForm
    {
        private const string Texto = "Duplicidade!, Esta Chave já existe no Banco de Dados.";
        string usuarioLogado = dados.Usuario;

        public frmProtocolo()
        {
            InitializeComponent();
            VerifyVersion(webBrowser);
            webBrowser.Navigate("https://www.nfe.fazenda.gov.br/portal/manifestacaoDestinatario.aspx?tipoConteudo=z+xSrb/veGA=");
            webBrowser.ScriptErrorsSuppressed = true;
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
                    fbeKey.SetValue(appName, 11001, RegistryValueKind.DWord);
                }


                //For 32 bit Machine 
                fbeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                if (fbeKey == null)
                    fbeKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION");
                using (fbeKey)
                {
                    fbeKey.SetValue(appName, 11001, RegistryValueKind.DWord);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(appName + "\n" + ex.ToString(), "Unexpected error setting browser mode!");
            }
        }
        public string Mensagem { get; private set; }
        private void webBrowser_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser.Document.GetElementById("ctl00_ContentPlaceHolder1_rbtSemChave").InvokeMember("click");
            HtmlElementCollection elc = this.webBrowser.Document.GetElementsByTagName("select");
            foreach (HtmlElement el in elc)
            {
                if (el.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_ddlEmissor"))
                {
                    el.GetElementsByTagName("option")[1].SetAttribute("selected", "selected");
                }
            }
        }
        private void ConsultaChaveSefaz()
        {
            try
            {
                int countg = dataGridViewSefaz.RowCount;
                label6.Text = countg.ToString();
                HtmlElementCollection Pesquisa = this.webBrowser.Document.GetElementsByTagName("span");
                foreach (HtmlElement Funcao in Pesquisa)
                {
                    if (Funcao.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_lblResultado"))
                    {
                        textBox4.Text = Funcao.GetAttribute("innerText");
                        if (textBox4.Text.Contains("Continue pesquisando"))
                        {
                            var links = webBrowser.Document.GetElementsByTagName("input");
                            foreach (HtmlElement link in links)
                            {
                                txtNSU.Text = link.GetAttribute("id").Replace("ctl00_ContentPlaceHolder1_rbtSemChave", "").Replace("ctl00_ContentPlaceHolder1_rbtComChave", "").Replace("ctl00_ContentPlaceHolder1_iptCNPJBase", "").Replace("ctl00_ContentPlaceHolder1_iptCNPJ", "").Replace("ctl00_ContentPlaceHolder1_btnPesquisar", "").Replace("ctl00_ContentPlaceHolder1_btnPesquisar", "").Replace("hiddenInputToUpdateATBuffer_CommonToolkitScripts", "").Replace("__VIEWSTATE", "").Replace("__VIEWSTATEGENERATOR", "").Replace("__EVENTVALIDATION", "").Replace("ctl00_ContentPlaceHolder1_iptNSU", "");
                                txtChave.Text = link.GetAttribute("value").Replace("rbtComChave", "").Replace("rbtSemChave", "").Replace("Enviar Consulta", "").Replace("wEPDwULLTExOTEzNzk2NzEPFCsAAhUCAABkFgJmD2QWAgIDD2QWEAIJDw8WAh4EVGV4dAUPMjYsNjE0IGJpbGjDtWVzZGQCDQ8PFgIfAAUOMSw4MTggbWlsaMO1ZXNkZAIPDw8WAh4LTmF2aWdhdGVVcmwFFWluZm9Fc3RhdGlzdGljYXMuYXNweGRkAhMPDxYCHwEFOWh0dHBzOi8vd3d3MS5zcGVkLmZhemVuZGEuZ292LmJyL2xvZ2luL3NwZWQvc3BlZG5mZWFjZXNzb2RkAhcPDxYCHwEFMnBlcmd1bnRhc0ZyZXF1ZW50ZXMuYXNweD90aXBvQ29udGV1ZG89NDdGSW83Mno5OXM9ZGQCJw88KwARAwAPFgQeC18hRGF0YUJvdW5kZx4LXyFJdGVtQ291bnQCBWQBEBYAFgAWAAwUKwAAFgJmD2QWDmYPDxYCHgdWaXNpYmxlaGRkAgEPZBYCZg9kFgICAQ8PFgYeDUFsdGVybmF0ZVRleHQFK01hbmlmZXN0byBFbGV0csO0bmljbyBkZSBEb2N1bWVudG9zIEZpc2NhaXMeD0NvbW1hbmRBcmd1bWVudAUmaHR0cHM6Ly9kZmUtcG9ydGFsLnN2cnMucnMuZ292LmJyL01kZmUeCEltYWdlVXJsBR1", "").Replace("Pesquisar", "").Replace("SemAssinatura", "").Replace("4A2556B2", "").Replace("OK", "").Replace("/wEPDwULLTExOTEzNzk2NzEPFCsAAhUCAABkFgJmD2QWAgIDD2QWEAIJDw8WAh4EVGV4dAUPMjYsNjE0IGJpbGjDtWVzZGQCDQ8PFgIfAAUOMSw4MTggbWlsaMO1ZXNkZAIPDw8WAh4LTmF2aWdhdGVVcmwFFWluZm9Fc3RhdGlzdGljYXMuYXNweGRkAhMPDxYCHwEFOWh0dHBzOi8vd3d3MS5zcGVkLmZhemVuZGEuZ292LmJyL2xvZ2luL3NwZWQvc3BlZG5mZWFjZXNzb2RkAhcPDxYCHwEFMnBlcmd1bnRhc0ZyZXF1ZW50ZXMuYXNweD90aXBvQ29udGV1ZG89NDdGSW83Mno5OXM9ZGQCJw88KwARAwAPFgQeC18hRGF0YUJvdW5kZx4LXyFJdGVtQ291bnQCBWQBEBYAFgAWAAwUKwAAFgJmD2QWDmYPDxYCHgdWaXNpYmxlaGRkAgEPZBYCZg9kFgICAQ8PFgYeDUFsdGVybmF0ZVRleHQFK01hbmlmZXN0byBFbGV0csO0bmljbyBkZSBEb2N1bWVudG9zIEZpc2NhaXMeD0NvbW1hbmRBcmd1bWVudAUmaHR0cHM6Ly9kZmUtcG9ydGFsLnN2cnMucnMuZ292LmJyL01kZmUeCEltYWdlVXJsBR1+L2ltYWdlbnMvYmFubmVyX21kZmVfT2ZmLnBuZxYCHgV0aXRsZQUrTWFuaWZlc3RvIEVsZXRyw7RuaWNvIGRlIERvY3VtZW50b3MgRmlzY2Fpc2QCAg9kFgJmD2QWAgIBDw8WBh8FBSZDb25oZWNpbWVudG8gZGUgVHJhbnNwb3J0ZSBFbGV0csO0bmljbx8GBR1odHRwOi8vd3d3LmN0ZS5mYXplbmRhLmdvdi5ich8HBSR+L2ltYWdlbnMvYmFubmVyc19WaXNpdGVfQ1RlX09mZi5wbmcWAh8IBSZDb25oZWNpbWVudG8gZGUgVHJhbnNwb3J0ZSBFbGV0csO0bmljb2QCAw9kFgJmD2QWAgIBDw8WBh8FBSlTaXN0ZW1hIFDDumJsaWNvIGRlIEVzY3JpdHVyYcOnw6NvIEZpc2NhbB8GBSNodHRwOi8vd3d3MS5yZWNlaXRhLmZhemVuZGEuZ292LmJyLx8HBSV+L2ltYWdlbnMvYmFubmVyc19WaXNpdGVfU3BlZF9PZmYucG5nFgIfCAUpU2lzdGVtYSBQw7pibGljbyBkZSBFc2NyaXR1cmHDp8OjbyBGaXNjYWxkAgQPZBYCZg9kFgICAQ8PFgYfBQUqU3VwZXJpbnRlbmTDqm5jaWEgZGEgWm9uYSBGcmFuY2EgZGUgTWFuYXVzHwYFGmh0dHA6Ly93d3cuc3VmcmFtYS5nb3YuYnIvHwcFIH4vaW1hZ2Vucy9iYW5uZXJzX21hbmF1c19PZmYucG5nFgIfCAUqU3VwZXJpbnRlbmTDqm5jaWEgZGEgWm9uYSBGcmFuY2EgZGUgTWFuYXVzZAIFD2QWAmYPZBYCAgEPDxYGHwUFMlBvcnRhbCBOYWNpb25hbCBkbyBCaWxoZXRlIGRlIFBhc3NhZ2VtIEVsZXRyw7RuaWNvHwYFJWh0dHBzOi8vZGZlLXBvcnRhbC5zdnJzLnJzLmdvdi5ici9CcGUfBwUcfi9pbWFnZW5zL2Jhbm5lcl9icGVfT2ZmLnBuZxYCHwgFMlBvcnRhbCBOYWNpb25hbCBkbyBCaWxoZXRlIGRlIFBhc3NhZ2VtIEVsZXRyw7RuaWNvZAIGDw8WAh8EaGRkAjMPZBYCAgEPZBYCZg9kFgICCQ9kFgwCAw8QDxYCHgdDaGVja2VkaGRkZGQCBQ8QDxYCHwlnZGRkZAIHD2QWAmYPZBYEAgMPDxYCHwAFCDEyNzA2Mjg5ZGQCBw8PFgIeD1ZhbGlkYXRpb25Hcm91cAUIc2VtQ2hhdmVkZAIJD2QWBGYPZBYCAgMPD2QWCB4KT25LZXlQcmVzcwUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KR4Hb25Gb2N1cwUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KR4Gb25CbHVyBRltYXNjYXJhKHRoaXMsc29OdW1lcm9zNDQpHghvbkNoYW5nZQUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KWQCAQ9kFgICEw88KwARAwAPFgYfAmcfAwICHwRnZAEQFgICAwIEFgI8KwAFAQAWAh4KSGVhZGVyVGV4dAUOU2l0dWHDp8OjbyBORmU8KwAFAQAWAh8PBRRTaXR1YcOnw6NvIE1hbmlmZXN0LhYCZmYMFCsAABYCZg9kFgYCAQ9kFgpmD2QWAmYPFQMLNDY4MDA2NjM4OTQsMzUyMTAyNTQzODM1MDAwMDAxODk1NTAwMTAwMDA0MjIyMTEwMDAzNzExMzdNYXR1YWxpemFyQ2hhdmVTZWxlY2lvbmFkYSgnMzUyMTAyNTQzODM1MDAwMDAxODk1NTAwMTAwMDA0MjIyMTEwMDAzNzExMzcnLCcwJylkAgEPDxYCHwAFCzQ2ODAwNjYzODk0ZGQCAg9kFgJmDxUOCzQ2ODAwNjYzODk0CzQ2ODAwNjYzODk0CzQ2ODAwNjYzODk0CzQ2ODAwNjYzODk0LDM1MjEwMjU0MzgzNTAwMDAwMTg5NTUwMDEwMDAwNDIyMjExMDAwMzcxMTM3CzQ2ODAwNjYzODk0HkVRVUlQRSBJTkRVU1RSSUEgTUVDQU5JQ0EgTFREQRI1NC4zODMuNTAwLzAwMDEtODkMNTM1MDA0NzE3MTEzEzAxLzAyLzIwMjEgMDA6MDA6MDAGU2HDrWRhC1IkIDQuMzEyLDAwHCtOYlhFczdiZElxcytqaUE0YUJGbjNBMWNoUT0TMDEvMDIvMjAyMSAwODoxNDowOGQCAw8PFgIfAAUKQXV0b3JpemFkYWRkAgQPDxYCHwAFGlNlbSBNYW5pZmVzdGEmIzIzMTsmIzIyNztvZGQCAg9kFgpmD2QWAmYPFQMLNDY4MDA2OTIwMDQsMjcyMTAyMDUxNDc3NDgwMDAxMzU1NTAwNDAwMDAxNzE5MzE4NTAwOTI2NDVNYXR1YWxpemFyQ2hhdmVTZWxlY2lvbmFkYSgnMjcyMTAyMDUxNDc3NDgwMDAxMzU1NTAwNDAwMDAxNzE5MzE4NTAwOTI2NDUnLCcxJylkAgEPDxYCHwAFCzQ2ODAwNjkyMDA0ZGQCAg9kFgJmDxUOCzQ2ODAwNjkyMDA0CzQ2ODAwNjkyMDA0CzQ2ODAwNjkyMDA0CzQ2ODAwNjkyMDA0LDI3MjEwMjA1MTQ3NzQ4MDAwMTM1NTUwMDQwMDAwMTcxOTMxODUwMDkyNjQ1CzQ2ODAwNjkyMDA0IkFHUk9DQU5BIENPTSBFIFJFUFJFU0VOVEFDT0VTIExUREESMDUuMTQ3Ljc0OC8wMDAxLTM1CTI0MTAyMDQ2OBMwMS8wMi8yMDIxIDAwOjAwOjAwBlNhw61kYQ1SJCAxMzkuMDAwLDAwHGRnTmFZUXNJeXNZNkFFMFFFTXRrNGNPVjMvcz0TMDEvMDIvMjAyMSAwODoxNDowMWQCAw8PFgIfAAUKQXV0b3JpemFkYWRkAgQPDxYCHwAFGlNlbSBNYW5pZmVzdGEmIzIzMTsmIzIyNztvZGQCAw8PFgIfBGhkZAILDw8WAh8EZ2QWBgIDDw8WAh8KBQhzZW1DaGF2ZWRkAgUPEGRkFgFmZAINDw8WAh8KBQhzZW1DaGF2ZWRkAg8PDxYCHwoFCHNlbUNoYXZlZGQCNw8PFgIfAAUtUG9ydGFsIGRhIE5GLWUgMjAyMSAtIE5vdGEgRmlzY2FsIEVsZXRyw7RuaWNhZGQYBQUeX19Db250cm9sc1JlcXVpcmVQb3N0QmFja0tleV9fFgkFD2N0bDAwJGlidEJ1c2NhcgUpY3RsMDAkZ2R2TGlua3NEZXN0YXF1ZSRjdGwwMiRJbWFnZUJ1dHRvbjEFKWN0bDAwJGdkdkxpbmtzRGVzdGFxdWUkY3RsMDMkSW1hZ2VCdXR0b24xBSljdGwwMCRnZHZMaW5rc0Rlc3RhcXVlJGN0bDA0JEltYWdlQnV0dG9uMQUpY3RsMDAkZ2R2TGlua3NEZXN0YXF1ZSRjdGwwNSRJbWFnZUJ1dHRvbjEFKWN0bDAwJGdkdkxpbmtzRGVzdGFxdWUkY3RsMDYkSW1hZ2VCdXR0b24xBSVjdGwwMCRDb250ZW50UGxhY2VIb2xkZXIxJHJidENvbUNoYXZlBSVjdGwwMCRDb250ZW50UGxhY2VIb2xkZXIxJHJidENvbUNoYXZlBSVjdGwwMCRDb250ZW50UGxhY2VIb2xkZXIxJHJidFNlbUNoYXZlBSljdGwwMCRDb250ZW50UGxhY2VIb2xkZXIxJG10dk1hbmlmZXN0YWNhbw8PZAIBZAUuY3RsMDAkQ29udGVudFBsYWNlSG9sZGVyMSRnZHZSZXN1bHRhZG9QZXNxdWlzYQ88KwAMAQgCAWQFJmN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkbXR2SW5zY3JpY2FvDw9kZmQFFmN0bDAwJGdkdkxpbmtzRGVzdGFxdWUPPCsADAEIAgFkmJ1ICoGnMMdQ6XF+w66s2IS2tUs=", "").Replace("/wEPDwULLTExOTEzNzk2NzEPFCsAAhUCAABkFgJmD2QWAgIDD2QWEAIJDw8WAh4EVGV4dAUPMjYsNjE0IGJpbGjDtWVzZGQCDQ8PFgIfAAUOMSw4MTggbWlsaMO1ZXNkZAIPDw8WAh4LTmF2aWdhdGVVcmwFFWluZm9Fc3RhdGlzdGljYXMuYXNweGRkAhMPDxYCHwEFOWh0dHBzOi8vd3d3MS5zcGVkLmZhemVuZGEuZ292LmJyL2xvZ2luL3NwZWQvc3BlZG5mZWFjZXNzb2RkAhcPDxYCHwEFMnBlcmd1bnRhc0ZyZXF1ZW50ZXMuYXNweD90aXBvQ29udGV1ZG89NDdGSW83Mno5OXM9ZGQCJw88KwARAwAPFgQeC18hRGF0YUJvdW5kZx4LXyFJdGVtQ291bnQCBWQBEBYAFgAWAAwUKwAAFgJmD2QWDmYPDxYCHgdWaXNpYmxlaGRkAgEPZBYCZg9kFgICAQ8PFgYeDUFsdGVybmF0ZVRleHQFK01hbmlmZXN0byBFbGV0csO0bmljbyBkZSBEb2N1bWVudG9zIEZpc2NhaXMeD0NvbW1hbmRBcmd1bWVudAUmaHR0cHM6Ly9kZmUtcG9ydGFsLnN2cnMucnMuZ292LmJyL01kZmUeCEltYWdlVXJsBR1+L2ltYWdlbnMvYmFubmVyX21kZmVfT2ZmLnBuZxYCHgV0aXRsZQUrTWFuaWZlc3RvIEVsZXRyw7RuaWNvIGRlIERvY3VtZW50b3MgRmlzY2Fpc2QCAg9kFgJmD2QWAgIBDw8WBh8FBSZDb25oZWNpbWVudG8gZGUgVHJhbnNwb3J0ZSBFbGV0csO0bmljbx8GBR1odHRwOi8vd3d3LmN0ZS5mYXplbmRhLmdvdi5ich8HBSR+L2ltYWdlbnMvYmFubmVyc19WaXNpdGVfQ1RlX09mZi5wbmcWAh8IBSZDb25oZWNpbWVudG8gZGUgVHJhbnNwb3J0ZSBFbGV0csO0bmljb2QCAw9kFgJmD2QWAgIBDw8WBh8FBSlTaXN0ZW1hIFDDumJsaWNvIGRlIEVzY3JpdHVyYcOnw6NvIEZpc2NhbB8GBSNodHRwOi8vd3d3MS5yZWNlaXRhLmZhemVuZGEuZ292LmJyLx8HBSV+L2ltYWdlbnMvYmFubmVyc19WaXNpdGVfU3BlZF9PZmYucG5nFgIfCAUpU2lzdGVtYSBQw7pibGljbyBkZSBFc2NyaXR1cmHDp8OjbyBGaXNjYWxkAgQPZBYCZg9kFgICAQ8PFgYfBQUqU3VwZXJpbnRlbmTDqm5jaWEgZGEgWm9uYSBGcmFuY2EgZGUgTWFuYXVzHwYFGmh0dHA6Ly93d3cuc3VmcmFtYS5nb3YuYnIvHwcFIH4vaW1hZ2Vucy9iYW5uZXJzX21hbmF1c19PZmYucG5nFgIfCAUqU3VwZXJpbnRlbmTDqm5jaWEgZGEgWm9uYSBGcmFuY2EgZGUgTWFuYXVzZAIFD2QWAmYPZBYCAgEPDxYGHwUFMlBvcnRhbCBOYWNpb25hbCBkbyBCaWxoZXRlIGRlIFBhc3NhZ2VtIEVsZXRyw7RuaWNvHwYFJWh0dHBzOi8vZGZlLXBvcnRhbC5zdnJzLnJzLmdvdi5ici9CcGUfBwUcfi9pbWFnZW5zL2Jhbm5lcl9icGVfT2ZmLnBuZxYCHwgFMlBvcnRhbCBOYWNpb25hbCBkbyBCaWxoZXRlIGRlIFBhc3NhZ2VtIEVsZXRyw7RuaWNvZAIGDw8WAh8EaGRkAjMPZBYCAgEPZBYCZg9kFgICCQ9kFgwCAw8QDxYCHgdDaGVja2VkaGRkZGQCBQ8QDxYCHwlnZGRkZAIHD2QWAmYPZBYEAgMPDxYCHwAFCDEyNzA2Mjg5ZGQCBw8PFgIeD1ZhbGlkYXRpb25Hcm91cAUIc2VtQ2hhdmVkZAIJD2QWBGYPZBYCAgMPD2QWCB4KT25LZXlQcmVzcwUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KR4Hb25Gb2N1cwUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KR4Gb25CbHVyBRltYXNjYXJhKHRoaXMsc29OdW1lcm9zNDQpHghvbkNoYW5nZQUZbWFzY2FyYSh0aGlzLHNvTnVtZXJvczQ0KWQCAQ9kFgICEw88KwARAgEQFgAWABYADBQrAABkAgsPDxYCHwRoZBYIAgMPDxYCHwoFCHNlbUNoYXZlZGQCBQ8QZGQWAWZkAgsPDxYCHwBlZGQCDQ8PFgIfCgUIc2VtQ2hhdmVkZAIPDw8WAh8KBQhzZW1DaGF2ZWRkAjcPDxYCHwAFLVBvcnRhbCBkYSBORi1lIDIwMjEgLSBOb3RhIEZpc2NhbCBFbGV0csO0bmljYWRkGAUFHl9fQ29udHJvbHNSZXF1aXJlUG9zdEJhY2tLZXlfXxYJBQ9jdGwwMCRpYnRCdXNjYXIFKWN0bDAwJGdkdkxpbmtzRGVzdGFxdWUkY3RsMDIkSW1hZ2VCdXR0b24xBSljdGwwMCRnZHZMaW5rc0Rlc3RhcXVlJGN0bDAzJEltYWdlQnV0dG9uMQUpY3RsMDAkZ2R2TGlua3NEZXN0YXF1ZSRjdGwwNCRJbWFnZUJ1dHRvbjEFKWN0bDAwJGdkdkxpbmtzRGVzdGFxdWUkY3RsMDUkSW1hZ2VCdXR0b24xBSljdGwwMCRnZHZMaW5rc0Rlc3RhcXVlJGN0bDA2JEltYWdlQnV0dG9uMQUlY3RsMDAkQ29udGVudFBsYWNlSG9sZGVyMSRyYnRDb21DaGF2ZQUlY3RsMDAkQ29udGVudFBsYWNlSG9sZGVyMSRyYnRDb21DaGF2ZQUlY3RsMDAkQ29udGVudFBsYWNlSG9sZGVyMSRyYnRTZW1DaGF2ZQUpY3RsMDAkQ29udGVudFBsYWNlSG9sZGVyMSRtdHZNYW5pZmVzdGFjYW8PD2QCAWQFLmN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkZ2R2UmVzdWx0YWRvUGVzcXVpc2EPZ2QFJmN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkbXR2SW5zY3JpY2FvDw9kZmQFFmN0bDAwJGdkdkxpbmtzRGVzdGFxdWUPPCsADAEIAgFkYPU3MMaZI90uLMS40LUlBm+6ivA=", "").Replace("/wEdACCZQKK7vRSRx4cQr1vYEFpQBGqPhwrln+9PTG+a425C4DQnJFM/mcs/d0VTMpMTYCF7arWIoAH6PP2G1YCHzk2BKmBS14FN+WjES3+DgGSxUhsv/8cgfVlU0c0VJB1DUuyMESBPWoOaEyn28Ulnl8glx4q7/POFF+ZhCwckO4rNNnFk901JR2pNXVbNcOI28xDN+LwgF7/qA/18pRQPHtbkdD7Xc1i83uy2RcKT19Xt0oWB3QqcHrat9AQl847by603YNPHBQBMCYcIffKKQyuNZwz77K2NFsDMYVxy7Rz9ZFEVCsZbzESC0Io+Uh033heiqdGYNZxbmcrurRp/zRPzXH16ZoHTII0N4WNnQmi39uS7Ad/dJxX0AdeXtOePjhHuGGF9M+Mxvu0wxo38I3j9MDR3qY3QLwCuQSDCQWUiHtbLJWvmIhvpV/TH6ilkyr9jwSlajPz+juOASAFOSnLvxX4zJujwnr3S+vdsZIjtyPnsuBk53VJ5/QjVX776QEfIx36AVoOLbK7qJmlnnVXS5IDR4kz9/BfZYB0/OL+tTQeHAMAxfxedYdp4fmCh4H2eOO0tTX51je1QOUgbhAXiTo1Ol8i+hK4G4Vo+jsDQOY8aEJZjnoPmyZdMj0cTDps02TTrU70chmk+oYZUyBFDTddRiEgPoDnKgxp8GtbRSplNGXywhWM49pzD+ImmeSYRpsLF", "").Replace("/wEdABlPHdHyh9YUee8+6l/u3+lxBGqPhwrln+9PTG+a425C4DQnJFM/mcs/d0VTMpMTYCF7arWIoAH6PP2G1YCHzk2BKmBS14FN+WjES3+DgGSxUhsv/8cgfVlU0c0VJB1DUuyMESBPWoOaEyn28Ulnl8glx4q7/POFF+ZhCwckO4rNNnFk901JR2pNXVbNcOI28xDN+LwgF7/qA/18pRQPHtbkdD7Xc1i83uy2RcKT19Xt0oWB3QqcHrat9AQl847by603YNPHBQBMCYcIffKKQyuNZwz77K2NFsDMYVxy7Rz9ZFEVCsZbzESC0Io+Uh033heiqdGYNZxbmcrurRp/zRPz+ey4GTndUnn9CNVfvvpAR8jHfoBWg4tsruomaWedVdLkgNHiTP38F9lgHT84v61NB4cAwDF/F51h2nh+YKHgfZ447S1NfnWN7VA5SBuEBeJOjU6XyL6ErgbhWj6OwNA5jxoQlmOeg+bJl0yPRxMOmzTZNOtTvRyGaT6hhlTIEUNN11GISA+gOcqDGnwa1tFKX7829YhsmqHVuhDwHKme5MteXu4=", "");
                                if (string.IsNullOrWhiteSpace(txtChave.Text))
                                {

                                }
                                else
                                {
                                    if (string.IsNullOrWhiteSpace(txtNSU.Text))
                                    {

                                    }
                                    else
                                    {
                                        VerificarChaveNoBanco();
                                    }
                                }
                            }
                            HtmlElementCollection elc = this.webBrowser.Document.GetElementsByTagName("input");
                            foreach (HtmlElement el in elc)
                            {
                                if (el.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_btnPesquisar"))
                                {
                                    el.InvokeMember("Click");
                                }
                            }
                            textBox4.Text = "";
                        }//Fim do Continue pesquisando
                        if (textBox4.Text.Contains("Nenhum documento localizado para o destinatario"))
                        {
                            checkBox.Checked = false;
                            TempoEspera.Enabled = true;
                            textBox4.Text = "";
                        }//Fim do Nenhum documento localizado para o destinatario
                        if (string.IsNullOrEmpty(textBox4.Text))
                        {
                            HtmlElementCollection elc = this.webBrowser.Document.GetElementsByTagName("input");
                            foreach (HtmlElement el in elc)
                            {
                                if (el.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_btnPesquisar"))
                                {
                                    el.InvokeMember("Click");
                                }
                            }
                        }
                    }//Fim do If ID-ctl00_ContentPlaceHolder1_lblResultado
                }//Fim do <Span>
                HtmlElementCollection Funcai = this.webBrowser.Document.GetElementsByTagName("input");
                foreach (HtmlElement el in Funcai)
                {
                    if (el.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_btnPesquisar"))
                    {
                        el.InvokeMember("Click");
                    }
                }
            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message);
            }
            LoadDataGrid();
        }
        private void LimparTexts()
        {
            //Limpar Campos apos a inserção no banco de dados.
            txtChave.Text = "";
            txtNSU.Text = "";
        }
        private void VerificarChaveNoBanco()
        {
            try
            {
                //Seleção da tabela no Banco de Dados.
                MySqlCommand prompt = new MySqlCommand("SELECT COUNT(*) FROM tb_chave WHERE col_chave ='" + txtChave.Text + "' ", ConexaoDados.GetConnectionXML());
                //Executa o comando.
                prompt.ExecuteNonQuery();
                //Converte o resultado para números inteiros.
                int consultDB = Convert.ToInt32(prompt.ExecuteScalar());
                ConexaoDados.GetConnectionXML().Close();
                //Verifica se o resultado for maior que zero(0), a execução inicia a Menssagem de que já existe contas, caso contrario faz a inserção no Banco.
                if (consultDB > 0)
                {
                    lblChaveDuplicidade.Text = txtChave.Text + ".xml";
                    LimparTexts();
                }
                else
                {
                    try
                    {
                        InserirChave();
                    }
                    catch (Exception Err)
                    {
                        MessageBox.Show(Err.Message);
                    }
                }
            }
            catch (Exception Err)
            {
                MessageBox.Show("Erro: " + Err.Message);
            }
        }
        private void InserirChave()
        {
            try
            {
                MySqlCommand prompt_cmd = new MySqlCommand("INSERT INTO tb_chave (col_chave,col_nsu,status,col_Downl) VALUES ('" + txtChave.Text.Trim() + "','" + txtNSU.Text.Trim() + "','.','1')", ConexaoDados.GetConnectionXML());
                prompt_cmd.ExecuteNonQuery();
                ConexaoDados.GetConnectionXML().Close();
            }
            catch (MySqlException ErroR)
            {
                MessageBox.Show("Erro ao Inserir Chave. " + ErroR.Message);
            }
            catch (Exception err)
            {
                MessageBox.Show("Erro: " + err.Message);
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            ConsultaChaveSefaz();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Baixar XML's?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                FormDownloadXML frm_Main = new FormDownloadXML();
                frm_Main.Show();
                Close();
            }
        }
        private void frmProtocolo_Load(object sender, EventArgs e)
        {
            LoadDataGrid();
            LoadDataGridPress();
            txtUrl.Text = @"C:\ArquivosSAP\XML\";
            label7.Text = @"C:\ArquivosSAP\XML\";
            lblChaveDuplicidade.Visible = false;
        }
        private void LoadDataGrid()
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT TipoDoc.col_desc_NFe, TbChave.* FROM `tb_chave` AS TbChave LEFT JOIN tb_tipo_nfe AS TipoDoc ON TbChave.tpNF=TipoDoc.col_id WHERE status!='LANÇADA' ORDER BY  col_nsu DESC", ConexaoDados.GetConnectionXML());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridViewSefaz.DataSource = SS;
            ConexaoDados.GetConnectionXML().Close();
            int countg = dataGridViewSefaz.RowCount;
            label6.Text = countg.ToString();
            ProgBar.Maximum = countg;
        }

        private void LoadDataGridPress()
        {
            MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT * FROM `tb_chave` WHERE status IS NULL", ConexaoDados.GetConnectionXML());
            DataTable SS = new DataTable();
            ADAP.Fill(SS);
            dataGridViewRestante.DataSource = SS;
            ConexaoDados.GetConnectionXML().Close();
            int countg = dataGridViewRestante.RowCount;
            label11.Text = countg.ToString();

        }
        private void dataGridViewSefaz_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }
        private void carregarDataXML()
        {
            int countg = dataGridViewSefaz.RowCount;
            label6.Text = countg.ToString();
            int XMLCont = 0;
            int XMLContResult = 1;
            double TotaldeLinhas = countg;
            while (XMLCont <= countg)
            {
                ProgBar.Value = XMLCont;
                txtURLxml.Text = dataGridViewSefaz.Rows[XMLCont].Cells["Column2"].Value.ToString();
                using (DataSet ds = new DataSet())
                {
                    try
                    {
                        string xmlFilePath = txtUrl.Text + txtURLxml.Text + ".xml";
                        ds.ReadXml(xmlFilePath);
                        try
                        {
                            dtXML.DataSource = ds.Tables["emit"];
                            txtCNPJ.Text = dtXML.Rows[0].Cells["CNPJ"].Value.ToString().Trim();
                            txtEmpresa.Text = dtXML.Rows[0].Cells["xNome"].Value.ToString().Trim();
                        }
                        catch (Exception ErroEmissor)
                        {
                            MessageBox.Show("Emit: " + ErroEmissor.Message);
                        }
                        try
                        {
                            NFEGRID.DataSource = ds.Tables["ide"];
                            txtNFE.Text = NFEGRID.Rows[0].Cells["nNF"].Value.ToString().Trim();
                            txtserie.Text = NFEGRID.Rows[0].Cells["serie"].Value.ToString().Trim();
                            txtdate.Text = NFEGRID.Rows[0].Cells["dhEmi"].Value.ToString().Trim().Substring(0, 10);
                            tpNF.Text = NFEGRID.Rows[0].Cells["tpNF"].Value.ToString().Trim();
                            dateTimePicker1.Text = txtdate.Text;
                        }
                        catch (Exception ErroIdent)
                        {
                            MessageBox.Show("Ide: " + ErroIdent.Message);
                        }
                        try
                        {
                            DADOSGRID.DataSource = ds.Tables["infAdic"];
                            txtDados.Text = DADOSGRID.Rows[0].Cells["infCpl"].Value.ToString().Trim();
                        }
                        catch (Exception)
                        {
                            //MessageBox.Show("InfAdic: " + ErroInfAdc.Message);
                        }
                        try
                        {
                            dataGridViewProdutos.DataSource = ds.Tables["ICMSTot"];
                            vNF.Text = dataGridViewProdutos.Rows[0].Cells["vNF"].Value.ToString().Replace(".", ",");
                            vNF.Focus();
                            textValor1.Focus();
                        }
                        catch (Exception ErroEmissor)
                        {
                            MessageBox.Show("total: " + ErroEmissor.Message);
                        }
                        try
                        {
                            CHAVEGRID.DataSource = ds.Tables["infProt"];
                            txtChavedeAcesso.Text = CHAVEGRID.Rows[0].Cells["chNFe"].Value.ToString().Trim();
                            txtProtocolo.Text = CHAVEGRID.Rows[0].Cells["nProt"].Value.ToString().Trim();
                        }
                        catch (Exception ErroInfProt)
                        {
                            MessageBox.Show("InfProt: " + ErroInfProt.Message);
                            txtCNPJ.Text = "";
                            txtEmpresa.Text = "";
                            txtNFE.Text = "";
                            txtserie.Text = "";
                            txtdate.Text = "";
                            txtDados.Text = "";
                            txtChavedeAcesso.Text = "";
                            txtProtocolo.Text = "";
                            vNF.Text = "";
                            tpNF.Text = "";
                        }
                        try
                        {
                            MySqlCommand prompt_cmd = new MySqlCommand("UPDATE `tb_chave` SET empresa='" + txtEmpresa.Text.Trim() + "' ,n_nfe='" + txtNFE.Text.Trim() + "',emisao='" + txtdate.Text.Trim() + "', tpNF='" + tpNF.Text.Trim() + "', vNF='" + vNF.Text.Trim() + "' WHERE col_chave='" + txtURLxml.Text.Replace(".xml", "") + "'", ConexaoDados.GetConnectionXML());
                            prompt_cmd.ExecuteNonQuery();
                            ConexaoDados.GetConnectionXML().Close();
                        }
                        catch (MySqlException)
                        {
                            //MessageBox.Show("Não Inserido, no Banco de Dados!");
                        }
                        catch (Exception Err)
                        {
                            MessageBox.Show("Geral: " + Err.Message);
                        }
                    }
                    catch (FileNotFoundException)
                    {
                        //MessageBox.Show(Error.Message + " " + XMLCont.ToString());
                    }
                    catch (Exception ErMXL)
                    {
                        MessageBox.Show("Geral 2: " + ErMXL.Message);
                    }
                }
                File.Delete(@"C:\ArquivosSAP\XML\" + txtURLxml.Text + ".xml");
                XMLCont++;
                XMLContResult++;
                try
                {
                    double percentual = XMLCont / TotaldeLinhas * 100.0;
                    lblPorcentagem.Text = percentual.ToString().Substring(0, 3) + " %";
                }
                catch (Exception)
                {
                    //MessageBox.Show("Porcentagem Error: " + ErrorR.Message);
                }
                label10.Text = XMLCont.ToString();
            }
            LoadDataGrid();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Iniciar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    carregarDataXML();
                }
                catch
                {
                    //MessageBox.Show("Geral 3: " + Err.Message);
                }
                LoadDataGrid();
            }
        }
        //richTextBox1.AppendText(el.GetAttribute("type").Equals("span") + Environment.NewLine);
        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Voltar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frm_Main frm_Main = new frm_Main();
                frm_Main.Show();
                this.Visible = false;
            }
        }
        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox.Checked)
            {
                TempoPesquisa.Enabled = true;
            }
            else
            {
                TempoPesquisa.Enabled = false;
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Iniciar a consulta no SAP? \nO SistemaGSG ficará indisponivel por alguns minutos", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //Pega a tela de execução do Windows
                CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
                //Pega a entrada ROT para o SAP Gui para conectar-se
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
                Session.SendCommand("/NJ1BNFE");
                int countg = dataGridViewSefaz.RowCount;

                //Tecla Enter
                guiWindow.SendVKey(0);


                ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtDATE0-LOW")).Text = "";
                string EMPRESA;
                string NOTAFISCAL;
                string Lancamento;
                string Protocolo;
                string User;
                string AcaoEtapa;
                string ValorNFe;
                int Chave = 0;
                int ProgressBarra = 1;
                while (Chave < countg)
                {
                    int countgXML = dataGridViewRestante.RowCount;
                    if (dataGridViewSefaz.Rows[Chave].Cells["Column10"].Value.ToString() == "LANÇADA")
                    {

                    }
                    else
                    {
                        try
                        {
                            ProgBar.Value = ProgressBarra;
                            ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtBUKRS-LOW")).Text = "USGA";
                            ((GuiTextField)Session.FindById("wnd[0]/usr/txtR_ACCKEY-LOW")).Text = dataGridViewSefaz.Rows[Chave].Cells["Column2"].Value.ToString();
                            ((GuiButton)Session.FindById("wnd[0]/tbar[1]/btn[8]")).Press();
                            AcaoEtapa = ((GuiGridView)Session.FindById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")).GetCellValue(0, "ACTION_REQU");
                            ((GuiGridView)Session.FindById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")).CurrentCellColumn = "";
                            ((GuiGridView)Session.FindById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")).SelectedRows = "0";
                            ((GuiGridView)Session.FindById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")).CurrentCellColumn = "NFNUM9";
                            ((GuiGridView)Session.FindById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")).ClickCurrentCell();
                            EMPRESA = ((GuiTextField)Session.FindById("wnd[0]/usr/subMAIN_PARTNER:SAPLJ1BB2:5250/txtJ_1BDYMPA-MAINNAME1")).Text;
                            NOTAFISCAL = ((GuiTextField)Session.FindById("wnd[0]/usr/subNF_NUMBER:SAPLJ1BB2:2002/txtJ_1BDYDOC-NFENUM")).Text;
                            Lancamento = ((GuiTextField)Session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-PSTDAT")).Text;
                            ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8")).Select();
                            Protocolo = ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/txtJ_1BDYDOC-AUTHCOD")).Text;
                            ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB7")).Select();
                            User = ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB7/ssubHEADER_TAB:SAPLJ1BB2:2700/txtJ_1BDYDOC-CRENAM")).Text;
                            ((GuiTab)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2")).Select();
                            ValorNFe = ((GuiTextField)Session.FindById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB2/ssubHEADER_TAB:SAPLJ1BB2:2200/txtJ_1BDYDOC-NFTOT")).Text;
                            ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                            ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                            if (AcaoEtapa == "C")
                            {
                                AcaoEtapa = "2";
                            }
                            if (string.IsNullOrEmpty(AcaoEtapa))
                            {
                                AcaoEtapa = "1";
                            }
                            if (AcaoEtapa == "A")
                            {
                                AcaoEtapa = "13";
                            }
                            if (AcaoEtapa == "B")
                            {
                                AcaoEtapa = "14";
                            }
                            if (AcaoEtapa == "D")
                            {
                                AcaoEtapa = "15";
                            }
                            dateTimePicker2.Text = Lancamento;
                            dateTimePicker2.Format = DateTimePickerFormat.Custom;
                            dateTimePicker2.CustomFormat = "yyyy-MM-dd";
                            try
                            {
                                MySqlCommand prompt_cmd = new MySqlCommand("UPDATE `tb_chave` SET empresa='" + EMPRESA + "' , n_nfe='" + NOTAFISCAL + "', lancamento_sap='" + this.dateTimePicker2.Text + "', protocolo='" + Protocolo + "',  user_sap='" + User + "', status='LANÇADA', vNF='R$: " + ValorNFe + "', col_Downl='2' ,  ACTION_REQU='" + AcaoEtapa + "' WHERE col_chave='" + dataGridViewSefaz.Rows[Chave].Cells["Column2"].Value.ToString() + "'", ConexaoDados.GetConnectionXML());
                                prompt_cmd.ExecuteNonQuery();
                                ConexaoDados.GetConnectionXML().Close();
                            }
                            catch (MySqlException ErrMy)
                            {
                                MessageBox.Show("Erro no Banco de Dados! - \n" + ErrMy.Message);
                            }
                            catch (Exception Err)
                            {
                                MessageBox.Show(Err.Message);
                            }
                        }
                        catch
                        {
                            try
                            {
                                MySqlCommand prompt_cmd = new MySqlCommand("UPDATE `tb_chave` SET status='NÃO LANÇADA' WHERE col_chave='" + dataGridViewSefaz.Rows[Chave].Cells["Column2"].Value.ToString() + "'", ConexaoDados.GetConnectionXML());
                                prompt_cmd.ExecuteNonQuery();
                                ConexaoDados.GetConnectionXML().Close();
                            }
                            catch (Exception er)
                            {
                                MessageBox.Show(er.Message);
                            }
                            ((GuiButton)Session.FindById("wnd[0]/tbar[0]/btn[3]")).Press();
                        }
                    }
                    label10.Text = Chave.ToString();
                    Chave++;
                    ProgressBarra++;
                    LoadDataGridPress();
                    try
                    {
                        double TotaldeLinhas = countg;
                        //double TotalLinhasNaoLanc = countgXML;
                        double percentual = Chave / TotaldeLinhas * 100.0;
                        double percent = 100.0 - percentual;
                        lblPorcentagem.Text = percentual.ToString().Substring(0, 3) + " %";
                    }
                    catch
                    {
                        // MessageBox.Show("Erro com a porcentagem!" +Error.Message);
                    }
                }
                LoadDataGrid();
            }
            MessageBox.Show("Processo, Finalizado com Sucesso!", "Conclusão", MessageBoxButtons.OK, MessageBoxIcon.None);
        }
        private void TempoEspera_Tick(object sender, EventArgs e)
        {
            HtmlElementCollection Botao = this.webBrowser.Document.GetElementsByTagName("input");
            foreach (HtmlElement Funcao in Botao)
            {
                if (Funcao.GetAttribute("id").Equals("ctl00_ContentPlaceHolder1_btnPesquisar"))
                {
                    Funcao.InvokeMember("Click");
                }
            }
            checkBox.Checked = true;
            TempoEspera.Enabled = false;
        }
        private void dataGridViewSefaz_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            var url = dataGridViewSefaz.SelectedRows[0].Cells["Column12"].Value;
            if (string.IsNullOrEmpty(url.ToString()))
            {

            }
            else
            {
                try
                {
                    var AbrirNavegador = new Navegador(url.ToString());
                    AbrirNavegador.Show();
                }
                catch (Exception Err)
                {
                    MessageBox.Show(Err.Message);
                }
            }
        }

        private void maskFiltro_MaskChanged(object sender, EventArgs e)
        {
        }

        private void maskFiltro_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string MES;
                string ANO;
                MES = maskFiltro.Text.Substring(0, 2);
                ANO = maskFiltro.Text.Substring(3, 4);

                MySqlDataAdapter ADAP = new MySqlDataAdapter("SELECT TipoDoc.col_desc_NFe, TbChave.* FROM `tb_chave` AS TbChave LEFT JOIN tb_tipo_nfe AS TipoDoc ON TbChave.tpNF=TipoDoc.col_id WHERE status!='LANÇADA' AND MONTH(emisao)='" + MES + "' AND  YEAR(emisao)='" + ANO + "' ORDER BY col_nsu DESC", ConexaoDados.GetConnectionXML());
                DataTable SS = new DataTable();
                ADAP.Fill(SS);
                dataGridViewSefaz.DataSource = SS;
                ConexaoDados.GetConnectionXML().Close();
                int countg = dataGridViewSefaz.RowCount;
                label6.Text = countg.ToString();
            }
            if (string.IsNullOrEmpty(maskFiltro.Text))
            {
                LoadDataGrid();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja ir para o Relatório Geral de Danf's?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                FormRelacao frm_Main = new FormRelacao();
                frm_Main.Show();
                Close();
            }
        }
    }
}
