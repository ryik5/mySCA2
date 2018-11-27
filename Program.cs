using System;
using System.Windows.Forms;

namespace mySCA2
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
            Application.Run(new FormPersonViewerSCA());
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
