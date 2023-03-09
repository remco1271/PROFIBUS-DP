using System;
using System.Windows.Forms;

namespace ProfDpTest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var profDpTest = new ProfDpTest())
            {
                Application.Run(profDpTest);
            }
        }
    }
}
