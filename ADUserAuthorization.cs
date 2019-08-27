using System;

namespace ASTA
{
    class ADUserAuthorization
    {
        public string Name { get; set; }       // имя
        public string Domain { get; set; }      // домен
        public string Password { get; set; }    // пароль
        public string DomainPath { get; set; }    // URI сервера

        public override string ToString()
        {
            return Name + "\t" + Domain + "\t" + Password + "\t" + DomainPath;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            ADUserAuthorization df = obj as ADUserAuthorization;
            if ((Object)df == null)
                return false;

            return this.ToString() == df.ToString();
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }

}
