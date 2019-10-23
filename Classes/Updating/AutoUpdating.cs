using System;

namespace ASTA.Classes.AutoUpdating
{
    public class AppUpdating
    {
        public delegate void Status(string message);
        public static event Status status;

        public static void Do(UpdatePreparing updating)
        {
            status?.Invoke("-= UpdatePreparing =-");
            updating.Do();
        }
    }

}
