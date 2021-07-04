using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SistemaGSG
{
    static class Program
    {
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            FormAcesso Sefaz = new FormAcesso();
            Sefaz.ShowDialog();

            //Splash fmr = new Splash();
            //fmr.ShowDialog();
            //frmLogin fml = new frmLogin();
            //fml.ShowDialog();
            //if (fml.FMP == true)
            //{ 
            //  Application.Run(new frm_Main());
            //}
        }
    }
}
