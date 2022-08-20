using System;
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

            //ConfigConexao frm_Main = new ConfigConexao();
            //frm_Main.ShowDialog();

            Splash fmr = new Splash();
            fmr.ShowDialog();
            frmLogin fml = new frmLogin();
            fml.ShowDialog();
            if (fml.FMP == true)
            {
              Application.Run(new frm_Main());
            }
        }
    }
}
