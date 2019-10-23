namespace ASTA.Classes.AutoUpdating
{    public class Updating
    {
        protected MakerOfLinks _subsystem1;
        protected MakerOfUpdateXmlFile _subsystem2;
        UpdatingParameters _parameters;

        public delegate void Status(object sender, AccountEventArgs e);
        public event Status status;

        public Updating(MakerOfLinks subsystem1, MakerOfUpdateXmlFile subsystem2, UpdatingParameters parameters)
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
            _subsystem2.MakeFile();

            status?.Invoke(this, new AccountEventArgs("Обновление завершено!"));
        }
    }
}
