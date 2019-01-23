using System;
using System.Windows.Forms;

namespace ASTA
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new WinFormASTA());
        }
    }

    static class UserLogin
    {
        public static string Value { get; set; }
    }

    static class UserPassword
    {
        public static string Value { get; set; }
    }

}
