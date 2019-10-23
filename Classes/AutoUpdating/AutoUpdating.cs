using System;

namespace ASTA.Classes.AutoUpdating
{
    public class MakerOfLinks
    {
        string _serverUpdateURL { get; set; }
        UpdatingParameters _parameters { get; set; }
        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;
        public MakerOfLinks(string serverUpdateURL)
        {
            _serverUpdateURL = serverUpdateURL;
        }

        public void Make()
        {
            if (string.IsNullOrWhiteSpace(_serverUpdateURL))
                throw new ArgumentNullException("Отсутствует параметр serverUpdateURL или ссылка пустая!", "serverUpdateURL");
            _parameters = new UpdatingParameters();

            _parameters.serverUpdateURL = _serverUpdateURL;
            _parameters.appUpdateFolderURL = @"file://" + _serverUpdateURL.Replace(@"\", @"/") + @"/";
            _parameters.appUpdateURL = _parameters.appUpdateFolderURL + _parameters.appFileXml;
            _parameters.appUpdateFolderURI = @"\\" + _serverUpdateURL + @"\";

            PrintProperties(_parameters);

            status?.Invoke(this, new AccountEventArgs("Все ссылки сгенерированы!"));
        }

        public UpdatingParameters GetParameters()
        {
            return _parameters;
        }

        private void PrintProperties(UpdatingParameters parameters)
        {
            foreach (var prop in parameters.GetType().GetProperties())
            {
                status?.Invoke(this, new AccountEventArgs(prop.Name + ": " + prop.GetValue(parameters, null)));
            }

            foreach (var field in parameters.GetType().GetFields())
            {
                status?.Invoke(this, new AccountEventArgs(field.Name + ": " + field.GetValue(parameters)));
            }
        }
    }

    
    public class AppUpdating
    {
        public delegate void Status(string message);
        public static event Status status;

        public static void ClientCode(Updating updating)
        {
            status?.Invoke("-= AppUpdating =-");
            updating.Do();
        }
    }

}
