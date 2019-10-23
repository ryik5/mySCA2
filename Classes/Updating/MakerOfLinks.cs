using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA.Classes.AutoUpdating
{
    public class MakerOfLinks
    {
        UpdatingParameters _parameters { get; set; }

        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;
        public MakerOfLinks(UpdatingParameters parameters)
        {
            _parameters = parameters;
        }

        public void SetParameters(UpdatingParameters parameters)
        {
            _parameters = parameters;
        }

        public void Make()
        {
            if (string.IsNullOrWhiteSpace(_parameters.remoteFolderUpdatingURL))
                throw new ArgumentNullException("Отсутствует параметр serverUpdateURL или ссылка пустая!", "serverUpdateURL");

            _parameters.remoteFolderUpdatingURL = _parameters.remoteFolderUpdatingURL;
            _parameters.appUpdateFolderURL = @"file://" + _parameters.remoteFolderUpdatingURL.Replace(@"\", @"/") + @"/";
            _parameters.appUpdateFolderURI = @"\\" + _parameters.remoteFolderUpdatingURL + @"\";
            _parameters.appUpdateURL = _parameters.appUpdateFolderURL + _parameters.appFileXml;

            PrintProperties(_parameters);

            status?.Invoke(this, new AccountEventArgs("Все ссылки сгенерированы!"));
        }

        public UpdatingParameters GetParameters()
        {
            return new UpdatingParameters {
                remoteFolderUpdatingURL = _parameters.remoteFolderUpdatingURL,
                localFolderUpdatingURL = _parameters.localFolderUpdatingURL,
                appUpdateFolderURL = _parameters.appUpdateFolderURL,
                appUpdateFolderURI = _parameters.appUpdateFolderURI,
                appUpdateURL = _parameters.appUpdateURL,
                appFileXml = _parameters.appFileXml,
                appUpdateChangeLogURL = _parameters.appUpdateChangeLogURL,
                appUpdateMD5 = _parameters.appUpdateMD5,
                appVersion = _parameters.appVersion,
                appFileZip = _parameters.appFileZip                
            };
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

}
