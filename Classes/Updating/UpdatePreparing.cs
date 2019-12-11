namespace ASTA.Classes.Updating
{
    public class UpdatePreparing
    {
        private IMakeable _makerLinks;
        private IMakeable _makerXml;
        private UpdatingParameters _parameters { set; get; }

        public delegate void Status(object sender, TextEventArgs e);

        public event Status status;

        public UpdatePreparing(IMakeable makerLinks, IMakeable makerXml, UpdatingParameters parameters)
        {
            _makerLinks = makerLinks;
            _makerXml = makerXml;
            _parameters = new UpdatingParameters(parameters);
        }

        public void Do()
        {
            _makerLinks.Set(_parameters);
            _makerLinks.Make();

            _parameters = _makerLinks.Get();

            _makerXml.Set(_parameters);
            _makerXml.Make();
            _parameters = _makerXml.Get();

            status?.Invoke(this, new TextEventArgs("Обновление для отправки на сервер подготовлено..."));
        }

        public UpdatingParameters GetParameters()
        {
            return new UpdatingParameters(_parameters);
        }
    }
}