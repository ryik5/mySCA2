using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASTA
{
    class MailServer
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

            MailServer df = obj as MailServer;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

    /* class MailServerBuilder
     {
         private MailServer _mailServer;

         public MailServerBuilder()
         { _mailServer = new MailServer(); }

         public MailServerBuilder(string name, int port)
         {
             _mailServer = new MailServer(name,port);
         }
         public void Set(string name, int port)
         {
             _mailServer = new MailServer(name, port);
         }
         public void SetName(string name)
         {
             _mailServer.SetName(name);
         }
         public void SetPort(int port)
         {
             _mailServer.SetPort(port);
         }

         public MailServer Build()
         {
             return _mailServer;
         }

         public override string ToString()
         {
             return _mailServer.ToString();
         }
     }*/

}
