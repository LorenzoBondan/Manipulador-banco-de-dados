using System;
using System.IO;
using System.Windows.Forms;

namespace ManipulacaoBanco
{
    internal static class Program
    {
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //if (!File.Exists(@"\\paris\eng\Usuarios\Lorenzo\BancoCaminho.sdf"))
            {
                //try
                {
                    //Application.Exit();
                    //return;
                }
                //catch (Exception)
                {
                    //Application.Exit();
                }
            }
            //else
            {
                Application.Run(new frmPrincipal());
            }
        }
    }
}
