using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SistemaGSG
{
    public partial class Navegador : MetroFramework.Forms.MetroForm
    {
        double milisegundos = 0;

        public Navegador(string UrlExec)
        {
            InitializeComponent();
            VerifyVersion(Browser);
            Browser.Navigate(UrlExec);
            timerFechar.Start();
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
        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void timerFechar_Tick(object sender, EventArgs e)
        {
            milisegundos += 15650;

            if (milisegundos >= 5000)
            {
                timerFechar.Stop();
                Close();
            }
        }
    }
}
