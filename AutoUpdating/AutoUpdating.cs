using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.AutoUpdating
{
    public class Updating
    {
        protected MakerOfLinks _subsystem1;
        protected MakerOfUpdateAppXML _subsystem2;
        UpdatingParameters _parameters;

        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;
        
        public Updating( MakerOfLinks subsystem1, MakerOfUpdateAppXML subsystem2, UpdatingParameters parameters)
        {
            _subsystem1 = subsystem1;
            _subsystem2 = subsystem2;
            _parameters = parameters;
        }

        public void Do()
        {
            _subsystem1.Make();
            _parameters = _subsystem1.GetParameters();
            _subsystem2.SetParameters(_parameters);
            _subsystem2.Make();

            status?.Invoke(this, new AccountEventArgs("Обновление завершено!"));
        }
    }

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

    public class UpdatingParameters
    {
        public string appUpdateMD5 { get; set; }
        public string appFileXml { get; set; }
        public string appVersion { get; set; }

        public string serverUpdateURL { get; set; }
        public string appUpdateFolderURL { get; set; }
        public string appUpdateURL { get; set; }
        public string appUpdateFolderURI { get; set; }
        public string appUpdateChangeLogURL { get; set; }
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
