﻿using System;
using System.Windows.Forms;

//For analysing crashes
using Microsoft.AppCenter;
using Microsoft.AppCenter.Analytics;
using Microsoft.AppCenter.Crashes;

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

            //For analysing crashes
            AppCenter.Start("b038f212-9e6f-4d65-8e4c-1f7fe0f9551a", typeof(Analytics), typeof(Crashes));
            
            //turn on gathering logs
            NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

            // get GIUD application
            string appGuid =
            ((System.Runtime.InteropServices.GuidAttribute)System.Reflection.Assembly.GetExecutingAssembly().
            GetCustomAttributes(typeof(System.Runtime.InteropServices.GuidAttribute), false).GetValue(0)).Value.ToString();

            logger.Info("");
            logger.Info("-------------------------------");
            logger.Info("");
            logger.Info("");
            logger.Info("-= Загрузка ПО =-");
            logger.Info("");

            using (System.Threading.Mutex mutex = new System.Threading.Mutex(false, "Global\\" + appGuid))
            {
                if (!mutex.WaitOne(0, false))
                {

                    //writing info about attempt to run another copy of the application
                    logger.Warn("Попытка запуска второй копии программы");
                    System.Threading.Tasks.Task.Run(() => MessageBox.Show("Программа уже запущена. Попытка запуска второй копии программы"));
                    System.Threading.Thread.Sleep(5000);
                    return;
                }

                //Running jnly one copy. Try to run the main application's form
                Application.Run(new WinFormASTA());
            }
        }
    }
}
