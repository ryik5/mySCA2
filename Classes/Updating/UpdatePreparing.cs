namespace ASTA.Classes.AutoUpdating
{    public class UpdatePreparing
    {
        protected MakerOfLinks _makerLinks;
        protected MakerOfUpdateXmlFile _makerXml;
      public UpdatingParameters _parameters { private set;   get ; }

        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;

        public UpdatePreparing(MakerOfLinks makerLinks, MakerOfUpdateXmlFile makerXml, UpdatingParameters parameters)
        {
            _makerLinks = makerLinks;
            _makerXml = makerXml;
            _parameters = parameters;
        }

        public void Do()
        {
             _makerLinks.SetParameters(_parameters);
           _makerLinks.Make();
            status?.Invoke(this, new AccountEventArgs("Все ссылки для выполнения операций обновления подготовлены."));

            _parameters = _makerLinks.GetParameters();

            _makerXml.SetParameters(_parameters);
            _makerXml.MakeFile();
            _parameters = _makerXml.GetParameters();
            status?.Invoke(this, new AccountEventArgs("Обновление для выгрузки на сервер подготовлены."));
        }

    }
}
