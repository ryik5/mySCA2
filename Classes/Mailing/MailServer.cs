namespace ASTA.Classes
{
    internal sealed class MailServer
    {
        private int _mailServerPort = -1;
        private string _mailServerName;

        public MailServer()
        { }

        public MailServer(string name, int port)
        {
            _mailServerPort = port;
            _mailServerName = name;
        }

        public void SetName(string name)
        {
            _mailServerName = name;
        }

        public void SetPort(int port)
        {
            _mailServerPort = port;
        }

        public string GetName()
        {
            return _mailServerName;
        }

        public int GetPort()
        {
            return _mailServerPort;
        }

        public override string ToString()
        {
            return _mailServerName + "\t" + _mailServerPort;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is MailServer))
                return false;

            var df = (MailServer)obj;
            if (df == null)
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}