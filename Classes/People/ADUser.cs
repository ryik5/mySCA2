namespace ASTA.Classes.People
{
    public interface IADUser
    {
        string Login { get; set; }        // имя
        string Domain { get; set; }       // домен
        string Password { get; set; }     // пароль
        string DomainControllerPath { get; set; }   // URI сервера
    }

    public class ADUser : IADUser
    {
        public string Login { get; set; }       // имя
        public string Domain { get; set; }      // домен
        public string Password { get; set; }    // пароль
        public string DomainControllerPath { get; set; }    // URI сервера

        public override string ToString()
        {
            return $"{Login}\t{Domain}\t{Password}\t{DomainControllerPath}";
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is ADUser df))
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}