using System.Diagnostics.Contracts;

namespace ASTA.Classes.Updating
{
    public class MakerOfLinks : IMakeable
    {
        private UpdatingParameters _parameters { get; set; }

        public delegate void Status(object sender, TextEventArgs e);

        public event Status status;

        public void Set(UpdatingParameters parameters)
        {
            Contract.Requires(parameters != null,
                 "Не создан экземпляр UpdatingParameters!");

            _parameters = parameters.Get();
        }

        public void Make()
        {
            Contract.Requires(_parameters != null,
                "Не создан экземпляр UpdatingParameters!");
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters.remoteFolderUpdatingURL),
                "Отсутствует параметр remoteFolderUpdatingURL или ссылка пустая!");
            Contract.Requires(!string.IsNullOrWhiteSpace(_parameters.appFileXml),
                "Отсутствует параметр appFileXml или ссылка пустая!");

            _parameters.appUpdateFolderURL = @"file://" + _parameters.remoteFolderUpdatingURL.Replace(@"\", @"/") + @"/";
            _parameters.appUpdateFolderURI = @"\\" + _parameters.remoteFolderUpdatingURL.Replace(@"/", @"\") + @"\";
            _parameters.appUpdateURL = _parameters.appUpdateFolderURL + _parameters.appFileXml;

            status?.Invoke(this, new TextEventArgs("Все ссылки сгенерированы!"));
        }

        public UpdatingParameters Get()
        {
            Contract.Requires(_parameters != null,
                "Не создан экземпляр UpdatingParameters!");
            
            return _parameters.Get();
        }

        /* public void PrintProperties()
         {
             foreach (var prop in _parameters.GetType().GetProperties())
             {
                 status?.Invoke(this, new TextEventArgs(prop.Name + ": " + prop.GetValue(_parameters, null)));
             }

             foreach (var field in _parameters.GetType().GetFields())
             {
                 status?.Invoke(this, new TextEventArgs(field.Name + ": " + field.GetValue(_parameters)));
             }
         }*/
    }
}