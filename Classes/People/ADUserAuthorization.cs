using System;

namespace ASTA.Classes.People
{
    class ADUserAuthorization
    {
        public string Login { get; set; }       // имя
        public string Domain { get; set; }      // домен
        public string Password { get; set; }    // пароль
        public string DomainPath { get; set; }    // URI сервера

        public override string ToString()
        {
            return Login + "\t" + Domain + "\t" + Password + "\t" + DomainPath;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is ADUserAuthorization df))
                return false;

            return ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}
