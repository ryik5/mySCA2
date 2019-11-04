﻿namespace ASTA.Classes.Updating
{
    public class UpdatePreparing
    {
        IMakeable _makerLinks;
        IMakeable _makerXml;
        UpdatingParameters _parameters { set; get; }

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
            _makerLinks.SetParameters(_parameters);
            _makerLinks.Make();

            _parameters = _makerLinks.GetParameters();

            _makerXml.SetParameters(_parameters);
            _makerXml.Make();
            _parameters = _makerXml.GetParameters();

            status?.Invoke(this, new TextEventArgs("Обновление для отправки на сервер подготовлено..."));
        }
        public UpdatingParameters GetParameters()
        {
            return new UpdatingParameters(_parameters);
        }
    }
}
